# Web-export

Web-export generates Word reports from the ViRVIG history workbook in your browser. For shared instructions, see the main README.

- [Generate a report in the browser](../readme.md#generate-word-reports-via-the-web-app)
- [Write a Jinja2 template and look up report fields](../readme.md#jinja2-in-word-templates)
- [Build, preview, and deploy the app](../readme.md#updating-the-web-export-app)

## Browser filters

Team and participant filters are optional for every report. Where year filters are available, you can leave either bound blank. In particular, journal and conference reports allow blank teams and years in the browser. Their CLI commands require a team and both year bounds.

## Built-in template catalog

The selector lists the built-in templates for the selected report type. The link below it downloads the selected built-in template. Choose **Use my own template…** to upload a DOCX instead. The build generates `dist/templates/catalogo.json` by scanning the repository's top-level `templates/` directory, with a subfolder for each report type. To add a built-in template, put the DOCX in its report's subfolder and rebuild. Its filename determines the label in the app.
