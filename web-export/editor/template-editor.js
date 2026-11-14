import { Editor, Node, mergeAttributes } from "@tiptap/core";
import { TableKit, TableRow } from "@tiptap/extension-table";
import TextAlign from "@tiptap/extension-text-align";
import StarterKit from "@tiptap/starter-kit";

const Field = Node.create({
  name: "field",
  group: "inline",
  inline: true,
  atom: true,
  addAttributes() {
    return { expression: { default: "" }, label: { default: "" },
      rich: { default: false }, condition: { default: "" } };
  },
  parseHTML() { return [{ tag: "span[data-template-field]" }]; },
  renderHTML({ node, HTMLAttributes }) {
    return ["span", mergeAttributes(HTMLAttributes, {
      "data-template-field": "", class: "field-chip", title: node.attrs.expression,
    }), node.attrs.label || node.attrs.expression];
  },
});

const Control = Node.create({
  name: "control",
  group: "block",
  atom: true,
  addAttributes() { return { action: { default: "end" }, value: { default: "" } }; },
  parseHTML() { return [{ tag: "div[data-template-control]" }]; },
  renderHTML({ node, HTMLAttributes }) {
    const { action, value } = node.attrs;
    const text = { repeat: `Repeat for each item in ${value}`, if: `If ${value}`, else: "Otherwise", end: "End block" };
    return ["div", mergeAttributes(HTMLAttributes, { "data-template-control": "", class: "template-control" }), text[action]];
  },
});

const RepeatRow = TableRow.extend({
  addAttributes() {
    return { ...this.parent?.(), repeat: {
      default: null,
      parseHTML: element => element.getAttribute("data-repeat"),
      renderHTML: ({ repeat }) => repeat ? { "data-repeat": repeat } : {},
    } };
  },
});

function repeatVariables(collection, catalog) {
  if (collection === "groups") return ["group", "loop"];
  return collection ? [catalog.item, "loop"] : [];
}

function scopeAtCursor(editor, catalog) {
  const { $from } = editor.state.selection;
  const stack = [];
  editor.state.doc.forEach((node, offset) => {
    if (offset >= $from.pos || node.type.name !== "control") return;
    const { action, value } = node.attrs;
    if (action === "repeat" || action === "if") stack.push(action === "repeat" ? value : null);
    if (action === "end") stack.pop();
  });
  for (let depth = $from.depth; depth > 0; depth--) {
    const node = $from.node(depth);
    if (node.type.name === "tableRow") stack.push(node.attrs.repeat);
  }
  return new Set(stack.flatMap(collection => repeatVariables(collection, catalog)));
}

function starter(catalog) {
  const field = expression => ({ type: "field", attrs: { expression, ...catalog.fields[expression] } });
  const text = value => ({ type: "text", text: value });
  return { type: "doc", content: [
    { type: "control", attrs: { action: "repeat", value: catalog.collection } },
    { type: "paragraph", content: [
      { ...field(catalog.item + ".title"), marks: [{ type: "bold" }] },
      ...(catalog.fields[catalog.item + ".citation"] ? [text(" · "), field(catalog.item + ".citation")] : []),
    ] },
    { type: "control", attrs: { action: "end" } },
    { type: "paragraph" },
  ] };
}

export function createTemplateEditor({ element, toolbar, catalog, content, onChange }) {
  const picker = toolbar.querySelector("[data-field-picker]");
  let fields = catalog;
  let previousScope;
  const editor = new Editor({
    element,
    content: content ?? starter(fields),
    extensions: [
      StarterKit.configure({ blockquote: false, codeBlock: false, horizontalRule: false,
        code: false, link: false, strike: false, heading: { levels: [1, 2, 3] } }),
      TableKit.configure({ table: { resizable: true }, tableRow: false }),
      RepeatRow, TextAlign.configure({ types: ["heading", "paragraph"] }), Field, Control,
    ],
    editorProps: { attributes: { role: "textbox", "aria-label": "Template editor", "aria-multiline": "true" } },
    onUpdate: ({ editor }) => { updateTools(editor); onChange(editor.getJSON()); },
    onSelectionUpdate: ({ editor }) => updateTools(editor),
  });

  function updateTools(instance) {
    const scope = scopeAtCursor(instance, fields);
    const key = [...scope].join(" ");
    if (key !== previousScope) {
      previousScope = key;
      picker.replaceChildren(new Option("Insert field…", ""), ...Object.entries(fields.fields)
        .filter(([expression]) => {
          const root = expression.split(".")[0];
          return !expression.includes(".") || scope.has(root);
        }).map(([expression, { label }]) => new Option(label, expression)));
    }
    for (const button of toolbar.querySelectorAll("[data-mark]")) {
      button.setAttribute("aria-pressed", String(instance.isActive(button.dataset.mark)));
    }
    for (const button of toolbar.querySelectorAll("[data-table]")) {
      button.disabled = !instance.isActive("table");
    }
  }

  function insertField(expression, attrs) {
    editor.chain().focus().insertContent({ type: "field", attrs: { expression, ...attrs } }).run();
  }

  function editExpression() {
    const selected = editor.isActive("field") ? editor.getAttributes("field") : {};
    const expression = window.prompt("Jinja expression without braces, for example pub.title | upper", selected.expression ?? "");
    if (!expression?.trim()) return;
    const condition = window.prompt("Optional condition for this field, for example pub.doi_url", selected.condition ?? "");
    if (condition === null) return;
    const attrs = { expression: expression.trim(), condition: condition.trim(),
      label: "", rich: fields.fields[expression.trim()]?.rich ?? false };
    if (selected.expression) editor.chain().focus().updateAttributes("field", attrs).run();
    else insertField(attrs.expression, attrs);
  }

  function insertControl(action, value = "") {
    const { $from } = editor.state.selection;
    const position = $from.depth ? $from.after(1) : $from.pos;
    editor.chain().focus().insertContentAt(position, { type: "control", attrs: { action, value } }).run();
  }

  function chooseCollection() {
    const scope = scopeAtCursor(editor, fields);
    const options = [fields.collection];
    if (fields.item === "pub") {
      options.push("groups");
      if (scope.has("group")) options.push("group.publications");
    }
    const choice = window.prompt(`Collection to repeat: ${options.join(", ")}`, scope.has("group") ? "group.publications" : fields.collection);
    return options.includes(choice?.trim()) ? choice.trim() : null;
  }

  const commands = {
    bold: () => editor.chain().focus().toggleBold().run(),
    italic: () => editor.chain().focus().toggleItalic().run(),
    underline: () => editor.chain().focus().toggleUnderline().run(),
    bullets: () => editor.chain().focus().toggleBulletList().run(),
    numbers: () => editor.chain().focus().toggleOrderedList().run(),
    expression: editExpression,
    repeat: () => { const key = chooseCollection(); if (key) insertControl("repeat", key); },
    condition: () => {
      const condition = window.prompt("Show the following blocks when this Jinja condition is true");
      if (condition?.trim()) insertControl("if", condition.trim());
    },
    otherwise: () => insertControl("else"),
    end: () => insertControl("end"),
    table: () => editor.chain().focus().insertTable({ rows: 2, cols: 3, withHeaderRow: true }).run(),
    row: () => editor.chain().focus().addRowAfter().run(),
    column: () => editor.chain().focus().addColumnAfter().run(),
    deleteRow: () => editor.chain().focus().deleteRow().run(),
    deleteColumn: () => editor.chain().focus().deleteColumn().run(),
    header: () => editor.chain().focus().toggleHeaderRow().run(),
    merge: () => editor.chain().focus().mergeOrSplit().run(),
    repeatRow: () => {
      const current = editor.getAttributes("tableRow").repeat;
      const key = current ? null : chooseCollection();
      if (current || key) {
        editor.chain().focus().updateAttributes("tableRow", { repeat: key }).run();
        updateTools(editor);
      }
    },
  };
  function onClick(event) {
    const command = event.target.closest("[data-command]")?.dataset.command;
    if (command) commands[command]();
  }
  function onSelect(event) {
    if (event.target === picker && picker.value) {
      insertField(picker.value, fields.fields[picker.value]);
      picker.value = "";
    } else if (event.target.matches("[data-text-style]")) {
      const value = Number(event.target.value);
      if (value) editor.chain().focus().setHeading({ level: value }).run();
      else editor.chain().focus().setParagraph().run();
    } else if (event.target.matches("[data-align]")) {
      editor.chain().focus().setTextAlign(event.target.value).run();
    }
  }
  toolbar.addEventListener("click", onClick);
  toolbar.addEventListener("change", onSelect);
  updateTools(editor);
  return {
    getJSON: () => editor.getJSON(),
    setEditable: enabled => editor.setEditable(enabled, false),
    setCatalog: catalog => { fields = catalog; previousScope = undefined; updateTools(editor); },
    destroy() {
      toolbar.removeEventListener("click", onClick);
      toolbar.removeEventListener("change", onSelect);
      editor.destroy();
    },
  };
}
