using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for Field and FieldType

class RemoveTocAndSaveHtml
{
    static void Main()
    {
        // Load the existing DOC/DOCX document.
        Document doc = new Document("input.docx");

        // Iterate over all fields in the document.
        // If a field is a Table of Contents (FieldToc), remove it from the document.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                // Remove the entire field node (including start, separator, and end).
                field.Remove();
                // No need to continue; there is typically only one TOC.
                break;
            }
        }

        // Prepare HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Ensure page numbers are not exported (default is false, set explicitly for clarity).
            ExportTocPageNumbers = false,
            // Omit headers and footers to keep the HTML clean (optional).
            ExportHeadersFootersMode = ExportHeadersFootersMode.None
        };

        // Save the modified document as HTML.
        doc.Save("output.html", saveOptions);
    }
}
