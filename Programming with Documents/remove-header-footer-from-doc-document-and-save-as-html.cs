using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting HTML file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.html";

        // Load the existing DOC document.
        Document doc = new Document(inputPath);

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Configure HTML save options to ensure headers/footers are not exported.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportHeadersFootersMode = ExportHeadersFootersMode.None
        };

        // Save the document as HTML.
        doc.Save(outputPath, saveOptions);
    }
}
