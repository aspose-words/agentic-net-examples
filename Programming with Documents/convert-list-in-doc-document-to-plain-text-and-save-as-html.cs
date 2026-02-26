using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextHtml
{
    static void Main()
    {
        // Path to the source DOC document that contains a list.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting HTML file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.html";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Ensure that list labels are up‑to‑date.
        doc.UpdateListLabels();

        // Configure HTML saving options to export list labels as plain inline text.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportListLabels = ExportListLabels.AsInlineText
        };

        // Save the document as HTML using the configured options.
        doc.Save(outputPath, htmlOptions);
    }
}
