const byId = (id) => document.getElementById(id);
const form = byId("export-form");
const fields = byId("form-fields");
const yearFilters = byId("year-filters");
const template = byId("template");
const templateDownload = byId("template-download");
const customTemplate = byId("custom-template");
const workbook = byId("workbook");
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
  byId("generate").disabled = false;
}

function updateTemplateSelection() {
  const custom = template.value === "custom";
  byId("custom-field").hidden = !custom;
  customTemplate.required = custom;
  customTemplate.disabled = !custom;
  const selected = catalog?.[selectedKind()]?.find(({ url }) => url === template.value);
  templateDownload.hidden = !selected;
  if (selected) {
    templateDownload.href = selected.url;
    templateDownload.textContent = `Download template '${selected.label}'`;
  } else {
    templateDownload.removeAttribute("href");
    templateDownload.textContent = "";
  }
}

function showStatus(message, state = "progress") {
  status.textContent = message;
  status.dataset.state = state;
  status.hidden = !message;
}

function setBusy(busy) {
  fields.disabled = busy;
  byId("generate").disabled = busy;
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
}

function stop(message, state) {
  worker?.terminate();
  worker = undefined;
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
  const engine = new Worker("worker.js");
  worker = engine;
  engine.onmessage = ({ data }) => {
    if (worker !== engine || !activeJob) return;
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
    if (worker === engine && activeJob) stop("Python could not run in the browser. Reload the page and try again.", "error");
  };
  return engine;
}

form.addEventListener("change", ({ target }) => {
  if (target.name === "kind") updateOptions();
  if (target === template) updateTemplateSelection();
});
form.addEventListener("input", () => {
  if (!activeJob) {
    clearDownload();
    showStatus("");
  }
});
cancel.addEventListener("click", () => stop("Export canceled."));

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (activeJob) return;
  clearDownload();
  const kind = selectedKind();
  const fromYear = yearFilters.disabled ? null : byId("from-year").valueAsNumber || null;
  const toYear = yearFilters.disabled ? null : byId("to-year").valueAsNumber || null;
  if (fromYear && toYear && fromYear > toYear) {
    showStatus("The start year must be on or before the end year.", "error");
    byId("from-year").focus();
    return;
  }
  const input = workbook.files[0];
  const custom = template.value === "custom";
  const word = customTemplate.files[0];
  if (!input || !/\.xlsx$/i.test(input.name) || (custom && (!word || !/\.docx$/i.test(word.name)))) {
    showStatus("Select an .xlsx workbook and, if you use your own template, a .docx file.", "error");
    return;
  }
  const teams = byId("teams").value;
  const participants = byId("participants").value;
  const selectedTemplate = template.value;
  const job = { kind, output: reportTypes[kind].output };
  activeJob = job;
  setBusy(true);
  showStatus("Preparing files…");
  startTimer(job, 150_000);
  try {
    const [workbookBytes, templateBytes] = await Promise.all([
      input.arrayBuffer(),
      custom ? word.arrayBuffer() : fetch(selectedTemplate, { cache: "no-cache" }).then((response) => {
        if (!response.ok) throw new Error("The built-in template could not be found.");
        return response.arrayBuffer();
      }),
    ]);
    if (activeJob !== job) return;
    getWorker().postMessage({
      kind, fromYear, toYear, teams, participants, workbook: workbookBytes, template: templateBytes,
    }, [workbookBytes, templateBytes]);
  } catch (error) {
    if (activeJob !== job) return;
    finish();
    showStatus(error.message, "error");
  }
});

updateOptions();
loadCatalog();
