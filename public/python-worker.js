/* IS — Image to PowerPoint browser worker */

const PYODIDE_VERSION = "0.29.5";
const PYODIDE_BASE = "https://cdn.jsdelivr.net/pyodide/v" + PYODIDE_VERSION + "/full/";
const PYODIDE_SCRIPT = PYODIDE_BASE + "pyodide.js";
const SOURCE_URL = "https://raw.githubusercontent.com/jipsy-danger/img-slide/main/build_slide.py";

let pyodide = null;
let booted = false;

function send(type, extra = {}) {
  self.postMessage({ type, ...extra });
}

function progress(percent, stage, detail = "") {
  send("progress", { percent, stage, detail });
}

async function ensureRuntime() {
  if (booted) return;

  progress(12, "Loading Python runtime", "Starting CPython through WebAssembly…");
  importScripts(PYODIDE_SCRIPT);
  pyodide = await loadPyodide({ indexURL: PYODIDE_BASE });

  progress(33, "Loading image engine", "Loading Pillow + XML support…");
  await pyodide.loadPackage(["lxml", "pillow"]);

  progress(49, "Loading PowerPoint engine", "Loading python-pptx…");
  await pyodide.runPythonAsync(
    "import micropip\\nawait micropip.install('python-pptx==1.0.2')"
  );

  pyodide.FS.mkdirTree("/workspace/input");
  pyodide.FS.mkdirTree("/mnt/user-data/outputs");
  booted = true;
  progress(61, "Runtime ready", "Browser Python engine is ready.");
}

async function fetchSource() {
  const response = await fetch(SOURCE_URL, { cache: "no-store" });
  if (!response.ok) {
    throw new Error("Could not load build_slide.py from GitHub.");
  }
  return await response.text();
}

function clearDirectory(path) {
  try {
    const entries = pyodide.FS.readdir(path);
    for (const name of entries) {
      if (name === "." || name === "..") continue;
      try { pyodide.FS.unlink(path + "/" + name); } catch (_) {}
    }
  } catch (_) {}
}

function locateOutput() {
  const entries = pyodide.FS.readdir("/mnt/user-data/outputs");
  const candidates = entries
    .filter((name) => /\.(pptx?|PPTX?)$/.test(name))
    .map((name) => "/mnt/user-data/outputs/" + name);

  if (!candidates.length) {
    throw new Error("The converter finished without creating a .pptx file.");
  }

  candidates.sort((a, b) => pyodide.FS.stat(b).size - pyodide.FS.stat(a).size);
  return candidates[0];
}

self.onmessage = async (event) => {
  if (event.data?.type !== "build") return;

  try {
    const incoming = Array.isArray(event.data.files) ? event.data.files : [];
    if (!incoming.length) {
      throw new Error("Select at least one image before starting conversion.");
    }

    await ensureRuntime();

    progress(66, "Preparing images", "Copying selected images into the Python input folder…");
    clearDirectory("/workspace/input");
    clearDirectory("/mnt/user-data/outputs");

    const usedNames = new Set();

    for (const item of incoming) {
      const original = String(item.name || "image");
      const safeBase = original.replaceAll("\\\\", "/").split("/").pop();
      if (!safeBase) continue;

      let safeName = safeBase;
      let n = 2;
      while (usedNames.has(safeName.toLowerCase())) {
        const dot = safeBase.lastIndexOf(".");
        const stem = dot > 0 ? safeBase.slice(0, dot) : safeBase;
        const ext = dot > 0 ? safeBase.slice(dot) : "";
        safeName = stem + "_" + n++ + ext;
      }
      usedNames.add(safeName.toLowerCase());

      pyodide.FS.writeFile(
        "/workspace/input/" + safeName,
        new Uint8Array(item.buffer)
      );
    }

    progress(75, "Running build_slide.py", "Converting each image into one PowerPoint slide…");

    const source = await fetchSource();
    pyodide.FS.writeFile("/workspace/build_slide.py", source);

    await pyodide.runPythonAsync(
      "import os\\n" +
      "os.environ['IMG_SLIDE_INPUT_DIR'] = '/workspace/input'\\n" +
      "os.environ['IMG_SLIDE_OUTPUT'] = '/mnt/user-data/outputs/IS_Images_to_PowerPoint.pptx'"
    );

    progress(82, "Building slides", "python-pptx is writing the presentation…");

    await pyodide.runPythonAsync(
      "exec(compile(open('/workspace/build_slide.py', 'r', encoding='utf-8').read(), " +
      "'build_slide.py', 'exec'), {'__name__': '__main__', '__file__': '/workspace/build_slide.py'})"
    );

    progress(93, "Finalizing output", "Reading the completed PowerPoint…");

    const outputPath = locateOutput();
    const outputName = outputPath.split("/").pop();
    const bytes = pyodide.FS.readFile(outputPath);

    progress(100, "Conversion complete", outputName + " is ready to download.");

    self.postMessage(
      {
        type: "done",
        name: outputName,
        mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        buffer: bytes.buffer,
      },
      [bytes.buffer]
    );
  } catch (error) {
    send("error", {
      message: error?.message ? String(error.message) : String(error),
      detail: error?.stack ? String(error.stack) : "",
    });
  }
};
