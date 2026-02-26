using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added namespace for Field and FieldType

class RemoveTocAndSaveHtml
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting HTML file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.html";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Find all Table of Contents (TOC) fields and remove them.
        // Collect the fields first to avoid modifying the collection while iterating.
        var tocFields = doc.Range.Fields
            .Where(f => f.Type == FieldType.FieldTOC)
            .ToList();

        foreach (Field toc in tocFields)
        {
            // Remove the entire field (start, separator, end) from the document.
            toc.Remove();
        }

        // Prepare HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Ensure that page numbers are not exported for any remaining TOC entries.
            ExportTocPageNumbers = false,

            // Optional: omit headers and footers if they are not needed.
            ExportHeadersFootersMode = ExportHeadersFootersMode.None
        };

        // Save the modified document as HTML.
        doc.Save(outputPath, saveOptions);
    }
}
