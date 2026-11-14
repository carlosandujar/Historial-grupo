# Web-export

Web-export generates Word reports from the ViRVIG history workbook in your browser. For shared instructions, see the main README.

- [Generate a report in the browser](../readme.md#generate-word-reports-via-the-web-app)
- [Write a Jinja2 template and look up report fields](../readme.md#jinja2-in-word-templates)
- [Build, preview, and deploy the app](../readme.md#updating-the-web-export-app)

## Browser filters

Team and participant filters are optional for every report. Where year filters are available, you can leave either bound blank. In particular, journal and conference reports allow blank teams and years in the browser. Their CLI commands require a team and both year bounds.

## Live preview

The read-only preview starts with fictional records. After you choose a workbook, it shows up to five records that match the current filters. It updates when you change the report type, filters, or template, including uploaded DOCX templates. The full report includes all matching records. Browser pagination can differ from Word.

## Visual template editor

Select a report type and choose **Create visual template**. The draft starts with a repeat block and a title field. Put the cursor inside the block to see the fields available there. For a table, place the cursor in a data row and choose **Repeat row** to insert record fields in that row. **Expression** accepts Jinja expressions without braces and an optional inline condition. Close each **Repeat block** or **If block** with **End block**.

The preview and full report use the active draft. **Download template** saves a reusable DOCX with its Jinja tags. You can use that file as a Word template through **Use my own template…**. Visual editing currently starts from a new draft.

To include the visual editor in a deployment, run `npm ci` and `npm run build` from `web-export`, alongside the [Python build](../readme.md#updating-the-web-export-app). The editor bundle is written to `dist/template-editor.js`.

## Built-in template catalog

The selector lists the built-in templates for the selected report type. The link below it downloads the selected built-in template. Choose **Use my own template…** to upload a DOCX instead. The build generates `dist/templates/catalogo.json` by scanning the repository's top-level `templates/` directory, with a subfolder for each report type. To add a built-in template, put the DOCX in its report's subfolder and rebuild. Its filename determines the label in the app.
