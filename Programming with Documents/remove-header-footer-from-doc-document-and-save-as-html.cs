using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveHtml
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting HTML file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.html";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Configure HTML save options to omit headers and footers.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportHeadersFootersMode = ExportHeadersFootersMode.None
        };

        // Save the document as HTML without headers and footers.
        doc.Save(outputPath, saveOptions);
    }
}
