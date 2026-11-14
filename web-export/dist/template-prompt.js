const examples = {
  title: "Terrain visualization", author: "Rivera, Alex", authors: "Rivera, Alex; Martin, Casey",
  supervisors: "Martin, Casey", contacts: "Rivera, Alex", team: "UPC", year: 2025,
  journal_title: "Example Journal", conference_proceedings: "Example Conference", journal_vol: 4,
  start_page: 12, end_page: 24, doi: "10.0000/example", doi_url: "https://doi.org/10.0000/example",
  title_with_link: "Terrain visualization (linked title)", doi_with_link: "https://doi.org/10.0000/example (linked)",
  citation: "Example Conference, pp. 12–24, 2025.", date: "01/03/2025", enrollment: "01/09/2025",
  description: "A public demonstration of terrain visualization.", activity_urls: "https://example.org/event",
  reference: "DEMO-2025", acronym: "DEMO", start: "01/01/2025", end: "31/12/2027",
  funding_entity: "Example funding agency", call: "Example call", ips: "Rivera, Alex", team_ip: "UPC",
  amount: 10000, amount_display: "10.000,00 €", text_es: "Proyecto de ejemplo.",
  links: ["https://example.org"], link_text: "Enlace 1 (linked)", has_links: true, active: true, eu: true, aei: false,
};

export function templatePrompt(kind, reportName, { collection, item, fields, yearLabel }) {
  const available = Object.entries(fields).map(([name, { label, rich }]) => {
    const key = name.slice(name.indexOf(".") + 1);
    const example = name === "loop.index" ? 1 : name.endsWith(" | length") ? 3
      : key === "citation" && kind === "journals" ? "Example Journal, vol. 4, pp. 12–24, 2025." : examples[key];
    return `- {{${rich ? "r" : ""} ${name} }} — ${label}${example !== undefined ? `; example: ${JSON.stringify(example)}` : ""}`;
  }).join("\n");
  return `Create a reusable Jinja2 Word report template for ViRVIG’s ${reportName} report. Reproduce the layout, language, labels, and punctuation in my example. Use fields for data instead of copying example values. Omit optional content when its value is blank. If no layout is supplied, ask me for one before generating the template.

Create and attach a downloadable .docx file containing the unrendered docxtpl template. Use real Word paragraphs, including one or more paragraphs per record. If my example uses a table, adapt it to labeled paragraphs. Use Word Normal and Heading 1–3 paragraph styles, with bold, italic, or underline applied to whole placeholders. Keep each Jinja tag intact in a single Word run. Use paragraph spacing and alignment for layout. Avoid tables, columns, page breaks, images, headers, footers, and custom character styles so the file can open in web-export’s visual editor.

Data and scope
Loop over ${collection} to access all matching records as ${item}. Use loop.index inside a repeat. To group by ${yearLabel}, repeat groups, use group.year as the heading, and repeat group.${collection} inside each group. A missing year is represented by None. Use ${item}.raw_cell with the exact Excel header for raw values.
Inside a record loop, you can repeat ${item}.display_fields as field to list its fields dynamically. Inside that loop use field.label and field.value, rendering the value as RichText only if field.rich is true.
The available fields are listed below. Example values are fictional. For fields without an example, use the field only when its meaning is clear.
${available}

Template rules
- Put each paragraph-level control tag in a paragraph of its own: {%p for ... %}, {%p if ... %}, {%p else %}, {%p endif %}, {%p endfor %}. Keep record fields inside their loop. For optional text within a paragraph, use {% if ... %} and {% endif %} in that same paragraph.
- Use {{ ${item}.title }} for plain fields and {{r ... }} only for listed RichText fields. Keep spaces inside delimiters as shown. For RichText, the exporter controls formatting and links; use the field directly without Jinja filters.
- Use simple Jinja expressions such as {{ ${item}.title | upper }}. Available filters include upper, lower, title, trim, length, and default. Keep function and method calls out of expressions.

Minimal example (one Word paragraph per line; adapt it to my layout)
{%p for ${item} in ${collection} %}
{{ loop.index }}. {{ ${item}.title }}
{%p endfor %}

My desired layout
[Replace this line with a representative example or a description of the desired layout. If you attach a screenshot, write “Use the attached screenshot.”]`;
}
