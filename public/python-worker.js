/* IMG-SLIDE browser Python worker
 * Runs the repository's build_slide.py inside Pyodide/WebAssembly.
 * User images are kept in the browser; no Python backend is required.
 */
const PYODIDE_VERSION = "0.29.5";
const PYODIDE_BASE = "https://cdn.jsdelivr.net/pyodide/v" + PYODIDE_VERSION + "/full/";
const PYODIDE_SCRIPT = PYODIDE_BASE + "pyodide.js";
const SOURCE_URL = "https://raw.githubusercontent.com/jipsy-danger/img-slide/main/build_slide.py";
const REQUIRED_ASSETS = [
  "logo.png", "p1g.png", "p2g.png", "cube.png", "center.png",
  "arrow.png", "grid.png", "th1.png", "th2.png", "ic1.png", "ic2.png", "ic3.png"
];

let pyodide = null;
let booted = false;

function send(type, extra = {}) {
  self.postMessage({ type, ...extra });
}

function setProgress(percent, stage, detail = "") {
  send("progress", { percent, stage, detail });
}

async function ensureRuntime() {
  if (booted) return;

  setProgress(12, "Loading Python runtime", "Starting CPython through WebAssembly…");

  importScripts(PYODIDE_SCRIPT);
  pyodide = await loadPyodide({ indexURL: PYODIDE_BASE });

  setProgress(34, "Loading PowerPoint engine", "Loading lxml + Pillow…");
  await pyodide.loadPackage(["lxml", "pillow"]);

  setProgress(49, "Loading python-pptx", "Installing the PowerPoint writer in the browser…");
  const micropip = pyodide.pyimport("micropip");
  await micropip.install("python-pptx==1.0.2");

  pyodide.FS.mkdirTree("/workspace");
  pyodide.FS.mkdirTree("/mnt/user-data/outputs");
  booted = true;

  setProgress(60, "Python engine ready", "Browser-side execution is ready.");
}

async function fetchSource() {
  const response = await fetch(SOURCE_URL, { cache: "no-store" });
  if (!response.ok) {
    throw new Error("Could not load build_slide.py from the repository.");
  }
  return await response.text();
}

function cleanOutputs() {
  try {
    const entries = pyodide.FS.readdir("/mnt/user-data/outputs");
    for (const name of entries) {
      if (name === "." || name === "..") continue;
      const path = "/mnt/user-data/outputs/" + name;
      try { pyodide.FS.unlink(path); } catch (_) {}
    }
  } catch (_) {}
}

function cleanWorkspace() {
  try {
    const entries = pyodide.FS.readdir("/workspace");
    for (const name of entries) {
      if (name === "." || name === "..") continue;
      const path = "/workspace/" + name;
      try { pyodide.FS.unlink(path); } catch (_) {}
    }
  } catch (_) {}
}

function findOutput() {
  const entries = pyodide.FS.readdir("/mnt/user-data/outputs");
  const candidates = entries
    .filter((name) => /\.(pptx?|PPTX?)$/.test(name))
    .map((name) => "/mnt/user-data/outputs/" + name);

  if (!candidates.length) {
    throw new Error("build_slide.py finished, but no PowerPoint file was created.");
  }
  candidates.sort((a, b) => pyodide.FS.stat(b).size - pyodide.FS.stat(a).size);
  return candidates[0];
}

self.onmessage = async (event) => {
  if (event.data?.type !== "build") return;

  try {
    const incoming = Array.isArray(event.data.files) ? event.data.files : [];
    const byName = new Map(incoming.map((item) => [String(item.name).toLowerCase(), item]));

    const missing = REQUIRED_ASSETS.filter(
      (name) => !byName.has(name.toLowerCase())
    );

    if (missing.length) {
      throw new Error(
        "Missing required assets: " + missing.join(", ") +
        ". The current build_slide.py uses these exact filenames."
      );
    }

    await ensureRuntime();

    setProgress(66, "Preparing assets", "Copying the uploaded image set into the Python workspace…");
    cleanWorkspace();
    cleanOutputs();

    for (const item of incoming) {
      const safeName = String(item.name).replaceAll("\\", "/").split("/").pop();
      if (!safeName) continue;
      pyodide.FS.writeFile("/workspace/" + safeName, new Uint8Array(item.buffer));
    }

    setProgress(76, "Executing build_slide.py", "The exact repository Python source is running now…");
    const source = await fetchSource();

    pyodide.FS.writeFile("/workspace/build_slide.py", source);
    pyodide.FS.chdir("/workspace");

    // Give build_slide.py its normal Python __main__ context.
    await pyodide.runPythonAsync(
      "exec(compile(open('/workspace/build_slide.py', 'r', encoding='utf-8').read(), " +
      "'build_slide.py', 'exec'), {'__name__': '__main__', '__file__': '/workspace/build_slide.py'})"
    );

    setProgress(91, "Finalizing PowerPoint", "Reading the generated .ppt/.pptx from the Python output folder…");
    const outputPath = findOutput();
    const outputName = outputPath.split("/").pop();
    const bytes = pyodide.FS.readFile(outputPath);

    setProgress(100, "Complete", outputName + " is ready.");
    self.postMessage(
      {
        type: "done",
        name: outputName,
        mime: outputName.toLowerCase().endsWith(".ppt")
          ? "application/vnd.ms-powerpoint"
          : "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        buffer: bytes.buffer
      },
      [bytes.buffer]
    );
  } catch (error) {
    send("error", {
      message: error?.message ? String(error.message) : String(error),
      detail: error?.stack ? String(error.stack) : ""
    });
  }
};
