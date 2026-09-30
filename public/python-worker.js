/* IS — minimal browser image → PowerPoint converter.
   build_slide.py is kept unchanged in the repository.
   This browser adapter converts uploaded images into PPTX slides. */

const PYODIDE_VERSION = "0.29.5";
const BASE = "https://cdn.jsdelivr.net/pyodide/v" + PYODIDE_VERSION + "/full/";
const SCRIPT_URL = BASE + "pyodide.js";

let pyodide = null;
let ready = false;

const send = (type, data = {}) => self.postMessage({ type, ...data });
const progress = (percent, stage, detail = "") => send("progress", { percent, stage, detail });

async function startRuntime() {
  if (ready) return;
  progress(8, "Starting", "Loading Python…");
  importScripts(SCRIPT_URL);
  pyodide = await loadPyodide({ indexURL: BASE });
  progress(25, "Preparing", "Loading image and PowerPoint packages…");
  await pyodide.loadPackage(["pillow", "lxml", "micropip"]);
  progress(42, "Preparing", "Loading python-pptx…");
  const micropip = pyodide.pyimport('micropip');
  await micropip.install('python-pptx==1.0.2');
  pyodide.FS.mkdirTree("/workspace/input");
  pyodide.FS.mkdirTree("/workspace/output");
  ready = true;
}

function clearDir(path) {
  try {
    for (const name of pyodide.FS.readdir(path)) {
      if (name === "." || name === "..") continue;
      try { pyodide.FS.unlink(path + "/" + name); } catch (_) {}
    }
  } catch (_) {}
}

self.onmessage = async (event) => {
  if (event.data?.type !== "build") return;
  try {
    const files = Array.isArray(event.data.files) ? event.data.files : [];
    if (!files.length) throw new Error("Select at least one image.");
    await startRuntime();
    progress(58, "Uploading", "Preparing " + files.length + " image" + (files.length === 1 ? "" : "s") + "…");
    clearDir("/workspace/input");
    clearDir("/workspace/output");

    for (let i = 0; i < files.length; i++) {
      const item = files[i];
      const name = String(item.name || ("image-" + (i + 1) + ".png")).replaceAll("\\", "/").split("/").pop();
      pyodide.FS.writeFile("/workspace/input/" + name, new Uint8Array(item.buffer));
    }

    progress(68, "Converting", "Creating one slide per image…");

    const python = [
      'from pathlib import Path',
      'import io',
      'from PIL import Image',
      'from pptx import Presentation',
      'from pptx.util import Emu',
      'from pptx.dml.color import RGBColor',
      '',
      'INPUT = Path("/workspace/input")',
      'OUTPUT = Path("/workspace/output/IS_Images_to_PowerPoint.pptx")',
      'SLIDE_W = Emu(12192000)',
      'SLIDE_H = Emu(6858000)',
      'EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".tif", ".tiff"}',
      '',
      'def add_image(prs, path):',
      '    slide = prs.slides.add_slide(prs.slide_layouts[6])',
      '    bg = slide.background.fill',
      '    bg.solid()',
      '    bg.fore_color.rgb = RGBColor(255, 255, 255)',
      '    with Image.open(path) as image:',
      '        image.load()',
      '        width_px, height_px = image.size',
      '        if width_px <= 0 or height_px <= 0:',
      '            raise ValueError(f"Invalid image: {path.name}")',
      '        image_bytes = io.BytesIO()',
      '        image.convert("RGBA").save(image_bytes, format="PNG")',
      '        image_bytes.seek(0)',
      '    sw, sh = int(SLIDE_W), int(SLIDE_H)',
      '    slide_ratio = sw / sh',
      '    image_ratio = width_px / height_px',
      '    if image_ratio >= slide_ratio:',
      '        width = sw',
      '        height = int(round(width / image_ratio))',
      '        left = 0',
      '        top = (sh - height) // 2',
      '    else:',
      '        height = sh',
      '        width = int(round(height * image_ratio))',
      '        top = 0',
      '        left = (sw - width) // 2',
      '    slide.shapes.add_picture(image_bytes, Emu(left), Emu(top), width=Emu(width), height=Emu(height))',
      '',
      'prs = Presentation()',
      'prs.slide_width = SLIDE_W',
      'prs.slide_height = SLIDE_H',
      'images = sorted([p for p in INPUT.iterdir() if p.is_file() and p.suffix.lower() in EXTS], key=lambda p: p.name.lower())',
      'if not images:',
      '    raise ValueError("No supported images found.")',
      'for image in images:',
      '    add_image(prs, image)',
      'prs.save(OUTPUT)',
    ].join("\n");

    await pyodide.runPythonAsync(python);

    progress(92, "Finishing", "Preparing the PowerPoint download…");
    const path = "/workspace/output/IS_Images_to_PowerPoint.pptx";
    const bytes = pyodide.FS.readFile(path);
    progress(100, "Done", "PowerPoint ready.");

    self.postMessage({
      type: "done",
      name: "IS_Images_to_PowerPoint.pptx",
      mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      buffer: bytes.buffer,
      slideCount: files.length
    }, [bytes.buffer]);
  } catch (error) {
    send('error', {
      message: error?.message ? String(error.message) : String(error),
      detail: error?.stack ? String(error.stack) : String(error)
    });
  }
};