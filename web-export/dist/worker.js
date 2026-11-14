const runtimeURL = "https://cdn.jsdelivr.net/pyodide/v0.29.3/full/";
const pythonFiles = ["export_publications.py", "export_conferences.py", "export_activities.py", "export_phd_theses.py", "export_phd_in_progress.py", "export_projects_international.py",
  "export_projects_national.py", "participant_filter.py", "workbook_cli.py", "report_workbook.py", "report_context.py",
  "sample_workbook.py", "template_fields.py", "template_docx.py", "template_import.py", "bridge.py"];
let runtime;

function reportError(error) {
  const message = String(error.message ?? error).trim().split("\n").at(-1);
  // Translate validation messages from the shared command-line exporters.
  const detail = message.replace(/^ValueError: /, "");
  return detail
    .replace(/^No se encontró la hoja (.+)\.$/, 'The workbook does not contain the "$1" worksheet.')
    .replace(/^(.+) debe tener una columna de cada tipo: (.+)\.$/, 'The "$1" worksheet must contain exactly one of each required column ($2).')
    .replace(/^(.+) tiene encabezados repetidos\. Cada columna con nombre debe ser única\.$/, 'The "$1" worksheet contains duplicate column headers. Each named column must be unique.')
    .replace(/^(.+) debe tener las columnas (.+) y (.+)\.$/, 'The "$1" worksheet must contain the $2 and $3 columns.')
    .replace(/^(.+) debe tener la columna Team o Team IP para filtrar por equipo\.$/, 'The "$1" worksheet needs a Team or Team IP column to filter by team.')
    .replace(/^(.+) debe tener la columna Team para filtrar por equipo\.$/, 'The "$1" worksheet needs a Team column to filter by team.')
    .replace(/^(.+) debe tener la columna IPs para filtrar por participantes\.$/, 'The "$1" worksheet needs an IPs column to filter by participants.')
    .replace(/^(.+) debe tener las columnas de fechas necesarias para filtrar por años\.$/, 'The "$1" worksheet needs the relevant date columns to filter by year.');
}

async function boot() {
  self.postMessage({ stage: "Loading Python in your browser…" });
  importScripts(`${runtimeURL}pyodide.js`);
  const python = await loadPyodide({ indexURL: runtimeURL });
  self.postMessage({ stage: "Loading Excel and Word libraries…" });
  await python.loadPackage(["micropip", "lxml", "jinja2"]);
  const pip = python.pyimport("micropip");
  try {
    await pip.install([
      "openpyxl==3.1.5", "docxtpl==0.20.1", "python-docx==1.2.0", "docxcompose==2.2.0",
    ]);
  } finally {
    pip.destroy();
  }
  await Promise.all(pythonFiles.map(async (name) => {
    const response = await fetch(name, { cache: "no-cache" });
    if (!response.ok) throw new Error(`Could not load ${name}.`);
    python.FS.writeFile(name, new Uint8Array(await response.arrayBuffer()));
  }));
  python.runPython("import bridge");
  return python;
}

self.onmessage = async ({ data }) => {
  let python;
  let bridge;
  try {
    runtime ??= boot();
    python = await runtime;
    if (data.editorRequest) {
      bridge = python.pyimport("bridge");
      if (data.editorRequest === "fields" || data.editorRequest === "field-reference") {
        self.postMessage({ editorRequest: data.editorRequest, kind: data.kind,
          request: data.request,
          catalog: JSON.parse(bridge.template_fields(data.kind)) });
      } else if (data.editorRequest === "open") {
        python.FS.writeFile("template-source.docx", new Uint8Array(data.baseTemplate));
        const draft = JSON.parse(bridge.inspect_template(data.kind));
        self.postMessage({ editorRequest: data.editorRequest, kind: data.kind, request: data.request,
          draft, catalog: JSON.parse(bridge.template_fields(data.kind)) });
      } else {
        if (data.baseTemplate) python.FS.writeFile("template-source.docx", new Uint8Array(data.baseTemplate));
        bridge.build_template.callKwargs(data.visualTemplate, data.baseTemplate ? "template-source.docx" : undefined,
          { embed: true });
        const bytes = python.FS.readFile("template.docx").slice();
        self.postMessage({ editorRequest: data.editorRequest, kind: data.kind, document: bytes.buffer }, [bytes.buffer]);
      }
      return;
    }
    if (data.inspect) {
      python.FS.writeFile("input.xlsx", new Uint8Array(data.workbook));
      bridge = python.pyimport("bridge");
      self.postMessage({ inspection: data.inspection, sheets: JSON.parse(bridge.load_history()) });
      return;
    }
    self.postMessage({ stage: data.preview ? "Rendering preview…" : "Generating the report…",
      rendering: true, preview: data.preview });
    if (data.workbook) {
      python.FS.writeFile("input.xlsx", new Uint8Array(data.workbook));
    }
    bridge = python.pyimport("bridge");
    if (data.visualTemplate) {
      if (data.baseTemplate) python.FS.writeFile("template-source.docx", new Uint8Array(data.baseTemplate));
      bridge.build_template(data.visualTemplate, data.baseTemplate ? "template-source.docx" : undefined);
    }
    else python.FS.writeFile("template.docx", new Uint8Array(data.template));
    if (data.workbook) bridge.load_history();
    // Pyodide maps undefined to None and null to JsNull.
    const counts = JSON.parse(bridge.generate.callKwargs(
      data.kind, data.fromYear ?? undefined, data.toYear ?? undefined, data.teams, data.participants,
      { preview: Boolean(data.preview), sample: Boolean(data.sample) },
    ));
    const document = python.FS.readFile("output.docx").slice();
    self.postMessage({ ...counts, preview: data.preview, document: document.buffer }, [document.buffer]);
  } catch (error) {
    self.postMessage({ error: reportError(error), setup: !python, inspection: data.inspection,
      preview: data.preview, editorRequest: data.editorRequest, kind: data.kind, request: data.request });
    if (!python) runtime = undefined;
  } finally {
    bridge?.destroy();
    for (const path of ["input.xlsx", "template-source.docx", "template.docx", "output.docx"]) {
      if (python?.FS.analyzePath(path).exists) python.FS.unlink(path);
    }
  }
};
