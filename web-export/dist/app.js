const byId = (id) => document.getElementById(id);
const form = byId("export-form");
const fields = byId("form-fields");
const yearFilters = byId("year-filters");
const template = byId("template");
const templateDownload = byId("template-download");
const customTemplate = byId("custom-template");
const workbook = byId("workbook");
const chooseWorkbook = byId("choose-workbook");
const download = byId("download");
const status = byId("status");
const cancel = byId("cancel");
const reportTypes = {
  journals: {
    sheet: "Journals", output: "publicaciones.docx", participantRole: "author",
  },
  conferences: {
    sheet: "Conferences", output: "conferencias.docx", participantRole: "author",
  },
  activities: {
    sheet: "Activities", output: "actividades.docx", participantRole: "contact",
    sheetHint: "Only rows marked Y, Yes, S, or Verdadero in the Export column are included.",
    countNouns: ["activity", "activities"],
  },
  theses: {
    sheet: "PhD Theses", output: "tesis_doctorales.docx",
  },
  progress: {
    sheet: "PhD in progress", output: "tesis_en_curso.docx",
  },
  projects: {
    sheet: "Projects-international", output: "proyectos_internacionales.docx", isProject: true,
  },
  national: {
    sheet: "Projects-national", output: "proyectos_nacionales.docx", isProject: true,
  },
};
let catalog;
let catalogFailed = false;
let worker;
let activeJob;
let downloadURL;
let workbookFile;
let workbookSheets;
let inspection = 0;
let inspectionTimer;
let workerWorkbook;
let previewJob;
let previewTimer;
let previewRevision = 0;
let previewPending = false;
let previewLibrary;
let visualEditor;
let visualKind;
let editorModule;
let editorDownloadPending = false;

function showEditorStatus(message, error = false) {
  const notice = byId("editor-status");
  notice.hidden = !message;
  notice.textContent = message;
  notice.classList.toggle("field-error", error);
}

function closeVisualEditor() {
  visualEditor?.destroy();
  visualEditor = undefined;
  visualKind = undefined;
  byId("visual-template").hidden = true;
  template.querySelector('option[value="visual"]')?.remove();
  showEditorStatus("");
}

function templatePayload() {
  return JSON.stringify({ kind: visualKind, document: visualEditor.getJSON() });
}

function requestEditorFields() {
  try {
    getWorker().postMessage({ editorRequest: "fields", kind: visualKind });
  } catch (error) {
    closeVisualEditor();
    showEditorStatus(error.message, true);
  }
}

async function openVisualEditor(data) {
  if (activeJob || data.kind !== visualKind || data.kind !== selectedKind()) return;
  if (data.error) {
    if (!visualEditor) closeVisualEditor();
    showEditorStatus(data.error, true);
    return;
  }
  if (visualEditor) {
    visualEditor.setCatalog(data.catalog);
    return;
  }
  try {
    editorModule ??= import("./template-editor.js?v=visual-editor").catch(error => {
      editorModule = undefined;
      throw error;
    });
    const { createTemplateEditor } = await editorModule;
    if (activeJob || data.kind !== visualKind || data.kind !== selectedKind() || visualEditor) return;
    visualEditor = createTemplateEditor({
      element: byId("template-editor"),
      toolbar: byId("editor-toolbar"),
      catalog: data.catalog,
      onChange: () => { clearDownload(); showStatus(""); schedulePreview(); },
    });
    template.add(new Option("Visual template draft", "visual"));
    template.value = "visual";
    updateTemplateSelection();
    clearDownload();
    showStatus("");
    showEditorStatus("");
    schedulePreview();
  } catch {
    closeVisualEditor();
    showEditorStatus("Could not load the visual editor. You can still use a Word template.", true);
  }
}

function reportFilters() {
  return {
    kind: selectedKind(),
    fromYear: yearFilters.disabled ? null : byId("from-year").valueAsNumber || null,
    toYear: yearFilters.disabled ? null : byId("to-year").valueAsNumber || null,
    teams: byId("teams").value,
    participants: byId("participants").value,
  };
}

async function readTemplate() {
  if (template.value === "visual") return { visualTemplate: templatePayload() };
  if (template.value === "custom") return { template: await customTemplate.files[0].arrayBuffer() };
  const response = await fetch(template.value, { cache: "no-cache", signal: AbortSignal.timeout(10_000) });
  if (!response.ok) throw new Error("The built-in template could not be found.");
  return { template: await response.arrayBuffer() };
}

function showPreviewStatus(message, error = false) {
  const notice = byId("preview-status");
  notice.textContent = message;
  notice.classList.toggle("field-error", error);
}

function schedulePreview() {
  previewRevision++;
  previewPending = true;
  clearTimeout(previewTimer);
  previewTimer = setTimeout(runPreview, 350);
}

async function runPreview() {
  if (!previewPending || activeJob || previewJob) return;
  previewPending = false;
  const filters = reportFilters();
  const sample = !workbookFile;
  const frame = byId("preview-document");
  const sheet = reportTypes[filters.kind].sheet;
  let waitMessage = "";
  if (template.disabled) waitMessage = "Loading templates…";
  else if (template.value === "custom" && !/\.docx$/i.test(customTemplate.files[0]?.name ?? "")) {
    waitMessage = "Choose a DOCX template to see a preview.";
  } else if (!sample && !workbookSheets?.includes(sheet)) {
    waitMessage = workbookSheets ? "Choose a workbook with the required worksheet." : "Waiting for a valid workbook…";
  } else if (!sample && filters.fromYear && filters.toYear && filters.fromYear > filters.toYear) {
    waitMessage = "The start year must be on or before the end year.";
  }
  if (waitMessage) {
    frame.hidden = true;
    frame.removeAttribute("srcdoc");
    showPreviewStatus(waitMessage);
    return;
  }

  const job = { revision: previewRevision, sample };
  previewJob = job;
  showPreviewStatus(frame.hidden ? "Preparing preview…" : "Updating preview…");
  job.timer = setTimeout(() => {
    if (previewJob !== job) return;
    if (job.usingWorker) resetWorker();
    finishPreview(job, "Preview timed out. Change a setting to retry.");
  }, 120_000);
  try {
    const source = await readTemplate();
    if (job !== previewJob || job.revision !== previewRevision || activeJob) return;
    const engine = getWorker();
    const workbookBytes = !sample && workerWorkbook !== workbookFile ? await workbookFile.arrayBuffer() : null;
    if (job !== previewJob || job.revision !== previewRevision || activeJob || engine !== worker) return;
    job.usingWorker = true;
    job.workbook = sample ? undefined : workbookFile;
    job.workerDone = new Promise(resolve => { job.releaseWorker = resolve; });
    engine.postMessage({ ...filters, ...source, preview: true, sample, workbook: workbookBytes },
      [source.template, workbookBytes].filter(Boolean));
  } catch (error) {
    finishPreview(job, error.message);
  } finally {
    if (previewJob === job && !job.usingWorker) finishPreview(job);
  }
}

function loadPreviewLibrary() {
  previewLibrary ??= (async () => {
    for (const url of [
      "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js",
      "https://cdn.jsdelivr.net/npm/docx-preview@0.3.6/dist/docx-preview.min.js",
    ]) {
      await new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.src = url;
        script.onload = resolve;
        script.onerror = () => {
          script.remove();
          reject(new Error("Could not load the Word preview viewer."));
        };
        document.head.append(script);
      });
    }
  })().catch(error => {
    previewLibrary = undefined;
    throw error;
  });
  return previewLibrary;
}

async function displayPreview(job, data) {
  try {
    if (job.revision !== previewRevision) return;
    if (data.error) throw new Error(data.error);
    await loadPreviewLibrary();
    const content = document.createElement("div");
    await globalThis.docx.renderAsync(data.document, content, null, {
      ignoreWidth: true, ignoreHeight: true, useBase64URL: true, renderAltChunks: false,
    });
    if (previewJob !== job || job.revision !== previewRevision) return;
    const frame = byId("preview-document");
    frame.srcdoc = `<!doctype html><meta charset="utf-8"><meta http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src 'unsafe-inline'; img-src data:; font-src data:">
      ${content.innerHTML}`;
    frame.hidden = false;
    showPreviewStatus(job.sample ? `Sample data · ${data.shown} example records`
      : data.count ? `Showing ${data.shown} of ${data.count} matching records`
      : "No records match the current filters.");
  } catch (error) {
    finishPreview(job, error.message);
  } finally {
    finishPreview(job);
  }
}

function finishPreview(job, error) {
  if (job !== previewJob) return;
  clearTimeout(job.timer);
  job.releaseWorker?.();
  previewJob = undefined;
  if (error && job.revision === previewRevision) {
    const previous = byId("preview-document").hidden ? "" : " Showing the previous preview.";
    showPreviewStatus(`Could not update the preview. ${error}${previous}`, true);
  }
  if (previewPending) {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(runPreview, 350);
  }
}

function selectedKind() {
  return form.elements.kind.value;
}

function updateOptions() {
  const kind = selectedKind();
  const { sheet, isProject, participantRole, sheetHint } = reportTypes[kind];
  yearFilters.hidden = yearFilters.disabled = kind === "progress";
  byId("project-years-hint").hidden = !isProject;
  if (isProject) yearFilters.setAttribute("aria-describedby", "project-years-hint");
  else yearFilters.removeAttribute("aria-describedby");
  byId("sheet-hint").textContent = [`This report uses the ${sheet} worksheet.`, sheetHint].filter(Boolean).join(" ");
  const role = participantRole ?? (isProject ? "principal investigator" : "thesis advisor");
  byId("participants-hint").textContent = `Enter ${role} last names separated by commas.`;
  if (catalog) {
    const options = catalog[kind] ?? [];
    template.replaceChildren(...options.map(({ label, url }) => new Option(label, url)),
      new Option("Use my own template…", "custom"));
    const message = byId("catalog-status");
    message.hidden = options.length > 0;
    message.textContent = options.length ? "" : catalogFailed
      ? "The template catalog could not be loaded. You can use your own DOCX template."
      : "No built-in templates are available for this report. You can use your own DOCX template.";
  }
  updateTemplateSelection();
  updateSheetStatus();
}

async function loadCatalog() {
  try {
    const response = await fetch("templates/catalogo.json", { cache: "no-cache", signal: AbortSignal.timeout(10_000) });
    if (!response.ok) throw new Error("Template catalog unavailable.");
    catalog = await response.json();
  } catch {
    catalog = {};
    catalogFailed = true;
  }
  updateOptions();
  template.disabled = false;
  updateSheetStatus();
  schedulePreview();
}

function updateTemplateSelection() {
  const custom = template.value === "custom";
  const visual = template.value === "visual";
  byId("visual-template").hidden = !visual;
  byId("custom-field").hidden = !custom;
  customTemplate.required = custom;
  customTemplate.disabled = !custom;
  const selected = catalog?.[selectedKind()]?.find(({ url }) => url === template.value);
  byId("template-name").textContent = visual ? "Visual template draft" : selected?.label ?? (custom
    ? customTemplate.files[0]?.name ?? "Choose a DOCX template."
    : "Loading templates…");
  templateDownload.hidden = !selected;
  if (selected) {
    templateDownload.href = selected.url;
    templateDownload.textContent = `Download template '${selected.label}'`;
  } else {
    templateDownload.removeAttribute("href");
    templateDownload.textContent = "";
  }
}

function updateSheetStatus() {
  const sheet = reportTypes[selectedKind()].sheet;
  const missing = workbookSheets && !workbookSheets.includes(sheet);
  const error = byId("sheet-error");
  error.hidden = !missing;
  error.textContent = missing ? `The workbook does not contain the “${sheet}” worksheet.` : "";
  byId("generate").disabled = Boolean(activeJob || template.disabled || !workbookSheets || missing);
}

function showWorkbookStatus(message, isError = false) {
  const notice = byId("workbook-status");
  notice.hidden = !message;
  notice.textContent = message;
  notice.classList.toggle("field-error", isError);
}

function showStatus(message, state = "progress") {
  status.textContent = message;
  status.dataset.state = state;
  status.hidden = !message;
}

function setBusy(busy) {
  fields.disabled = busy;
  chooseWorkbook.disabled = busy;
  byId("editor-toolbar").disabled = busy;
  byId("download-visual-template").disabled = busy || editorDownloadPending;
  visualEditor?.setEditable(!busy);
  updateSheetStatus();
  cancel.hidden = !busy;
}

function clearDownload() {
  if (downloadURL) URL.revokeObjectURL(downloadURL);
  downloadURL = undefined;
  download.removeAttribute("href");
  download.hidden = true;
}

function finish() {
  clearTimeout(activeJob?.timer);
  activeJob = undefined;
  setBusy(false);
  if (previewPending) schedulePreview();
}

function resetWorker() {
  worker?.terminate();
  worker = undefined;
  workerWorkbook = undefined;
  if (editorDownloadPending) {
    editorDownloadPending = false;
    byId("download-visual-template").disabled = Boolean(activeJob);
    showEditorStatus("Template download was interrupted. Try again.", true);
  }
}

function stop(message, state) {
  resetWorker();
  if (previewJob?.usingWorker) {
    finishPreview(previewJob);
    previewPending = true;
  }
  finish();
  showStatus(message, state);
}

function startTimer(job, milliseconds) {
  clearTimeout(job.timer);
  job.timer = setTimeout(() => {
    if (activeJob === job) stop("The export timed out. Check your connection or try another template.", "error");
  }, milliseconds);
}

function getWorker() {
  if (worker) return worker;
  const engine = new Worker("worker.js?v=visual-editor");
  worker = engine;
  engine.onmessage = ({ data }) => {
    if (worker !== engine) return;
    if (data.editorRequest === "fields") {
      openVisualEditor(data);
      return;
    }
    if (data.editorRequest === "download") {
      editorDownloadPending = false;
      byId("download-visual-template").disabled = Boolean(activeJob);
      if (data.error) showEditorStatus(data.error, true);
      else {
        const url = URL.createObjectURL(new Blob([data.document], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }));
        const link = document.createElement("a");
        link.href = url;
        link.download = `${data.kind ?? selectedKind()}-template.docx`;
        link.click();
        setTimeout(() => URL.revokeObjectURL(url), 10_000);
        showEditorStatus("");
      }
      return;
    }
    if (data.inspection !== undefined) {
      if (data.inspection !== inspection) return;
      clearTimeout(inspectionTimer);
      workbookSheets = data.sheets;
      workerWorkbook = data.sheets ? workbookFile : undefined;
      showWorkbookStatus(data.error
        ? (data.setup ? "Could not start Python. Check your connection and choose the workbook again." : `Could not read the workbook. ${data.error}`)
        : "", Boolean(data.error));
      updateSheetStatus();
      if (visualKind) requestEditorFields();
      schedulePreview();
      return;
    }
    if (data.preview) {
      if (!previewJob || data.stage) return;
      previewJob.usingWorker = false;
      previewJob.releaseWorker?.();
      if (!data.error && previewJob.workbook === workbookFile) workerWorkbook = previewJob.workbook;
      displayPreview(previewJob, data);
      return;
    }
    if (!activeJob) {
      if (data.stage && workbookFile && !workbookSheets) showWorkbookStatus(data.stage);
      return;
    }
    if (data.stage) {
      showStatus(data.stage);
      if (data.rendering) startTimer(activeJob, 45_000);
      return;
    }
    if (data.error) {
      if (data.setup) stop("Python could not start. Check your connection and try again.", "error");
      else {
        finish();
        showStatus(`Could not generate the report. ${data.error}`, "error");
      }
      return;
    }
    workerWorkbook = workbookFile;
    downloadURL = URL.createObjectURL(new Blob([data.document], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }));
    download.href = downloadURL;
    download.download = activeJob.output;
    download.hidden = false;
    const kind = activeJob.kind;
    const { isProject, countNouns } = reportTypes[kind];
    const [singular, plural] = countNouns ?? (kind === "progress" ? ["thesis in progress", "theses in progress"]
      : isProject ? ["project", "projects"] : ["reference", "references"]);
    const noun = data.count === 1 ? singular : plural;
    finish();
    showStatus(data.count ? `Report ready with ${data.count} ${noun}.` : "Report ready. No records match the selection.", "success");
  };
  engine.onerror = () => {
    if (worker !== engine) return;
    if (activeJob) {
      if (previewJob) finishPreview(previewJob, "Python could not run in the browser.");
      stop("Python could not run in the browser. Reload the page and try again.", "error");
    }
    else {
      resetWorker();
      if (visualKind && !visualEditor) {
        closeVisualEditor();
        showEditorStatus("Could not start the visual editor. Try again.", true);
      }
      if (previewJob) finishPreview(previewJob, "Python could not run in the browser.");
      if (workbookFile && !workbookSheets) {
        clearTimeout(inspectionTimer);
        showWorkbookStatus("Could not read the workbook. Choose it again to retry.", true);
      }
    }
  };
  return engine;
}

byId("create-template").addEventListener("click", () => {
  if (visualEditor && visualKind === selectedKind()) {
    template.value = "visual";
    updateTemplateSelection();
    clearDownload();
    showStatus("");
    schedulePreview();
    return;
  }
  visualKind = selectedKind();
  showEditorStatus("Loading visual editor…");
  requestEditorFields();
});

byId("download-visual-template").addEventListener("click", () => {
  if (editorDownloadPending || !visualEditor || activeJob) return;
  try {
    editorDownloadPending = true;
    byId("download-visual-template").disabled = true;
    showEditorStatus("Preparing template…");
    getWorker().postMessage({ editorRequest: "download", kind: visualKind, visualTemplate: templatePayload() });
  } catch (error) {
    editorDownloadPending = false;
    byId("download-visual-template").disabled = false;
    showEditorStatus(error.message, true);
  }
});

chooseWorkbook.addEventListener("click", () => workbook.click());
workbook.addEventListener("change", async () => {
  const file = workbook.files[0];
  workbook.value = "";
  if (!file) return;
  clearTimeout(inspectionTimer);
  const request = ++inspection;
  workbookFile = file;
  workbookSheets = undefined;
  byId("workbook-name").textContent = file.name;
  chooseWorkbook.textContent = "Replace workbook";
  clearDownload();
  showStatus("");
  updateSheetStatus();
  schedulePreview();
  if (!/\.xlsx$/i.test(file.name)) {
    showWorkbookStatus("Select an .xlsx workbook.", true);
    return;
  }
  showWorkbookStatus("Reading workbook…");
  inspectionTimer = setTimeout(() => {
    if (request !== inspection) return;
    inspection++;
    resetWorker();
    if (previewJob) finishPreview(previewJob);
    showWorkbookStatus("Reading the workbook timed out. Choose it again to retry.", true);
  }, 150_000);
  try {
    const bytes = await file.arrayBuffer();
    if (request !== inspection) return;
    getWorker().postMessage({ inspect: true, inspection: request, workbook: bytes }, [bytes]);
  } catch (error) {
    if (request === inspection) {
      clearTimeout(inspectionTimer);
      showWorkbookStatus(`Could not read the workbook. ${error.message}`, true);
    }
  }
});

form.addEventListener("change", ({ target }) => {
  if (target.name === "kind") {
    if (visualEditor && !window.confirm("Discard the visual template draft and switch report types?")) {
      form.querySelector(`input[name="kind"][value="${visualKind}"]`).checked = true;
      return;
    }
    closeVisualEditor();
    updateOptions();
  }
  if (target === template || target === customTemplate) updateTemplateSelection();
  if (fields.contains(target) && !activeJob) {
    clearDownload();
    showStatus("");
    schedulePreview();
  }
});
form.addEventListener("input", ({ target }) => {
  if (!activeJob && fields.contains(target) && target.type !== "file" && target.type !== "radio" && target !== template) {
    clearDownload();
    showStatus("");
    schedulePreview();
  }
});
cancel.addEventListener("click", () => stop("Export canceled."));

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (activeJob || template.disabled || !workbookSheets?.includes(reportTypes[selectedKind()].sheet)) return;
  clearDownload();
  const filters = reportFilters();
  const { kind, fromYear, toYear } = filters;
  if (fromYear && toYear && fromYear > toYear) {
    showStatus("The start year must be on or before the end year.", "error");
    byId("from-year").focus();
    return;
  }
  const input = workbookFile;
  const custom = template.value === "custom";
  const word = customTemplate.files[0];
  if (!input || !/\.xlsx$/i.test(input.name) || (custom && (!word || !/\.docx$/i.test(word.name)))) {
    showStatus("Select an .xlsx workbook and, if you use your own template, a .docx file.", "error");
    return;
  }
  const job = { kind, output: reportTypes[kind].output };
  activeJob = job;
  setBusy(true);
  showStatus("Preparing files…");
  startTimer(job, 150_000);
  try {
    // The worker processes one document at a time. A queued export takes
    // priority over the next automatic preview.
    if (previewJob?.usingWorker) await previewJob.workerDone;
    if (activeJob !== job) return;
    const [workbookBytes, source] = await Promise.all([
      workerWorkbook === input ? null : input.arrayBuffer(),
      readTemplate(),
    ]);
    if (activeJob !== job) return;
    getWorker().postMessage({
      ...filters, ...source, workbook: workbookBytes,
    }, [source.template, workbookBytes].filter(Boolean));
  } catch (error) {
    if (activeJob !== job) return;
    finish();
    showStatus(error.message, "error");
  }
});

updateOptions();
loadCatalog();
