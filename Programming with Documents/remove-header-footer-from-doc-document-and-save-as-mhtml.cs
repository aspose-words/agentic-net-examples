using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace RemoveHeadersFootersAndSaveMhtml
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOC document.
            string inputPath = @"C:\Docs\SourceDocument.doc";

            // Path where the resulting MHTML file will be saved.
            string outputPath = @"C:\Docs\ResultDocument.mht";

            // Load the existing DOC document.
            Document doc = new Document(inputPath);

            // Remove all header and footer contents from each section.
            // The HeaderFooter objects remain, but their text is cleared,
            // causing the document to have no visible headers/footers.
            foreach (Section section in doc.Sections)
            {
                section.ClearHeadersFooters();
            }

            // Configure save options for MHTML.
            // ExportHeadersFootersMode = None ensures that no header/footer data is written.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportHeadersFootersMode = ExportHeadersFootersMode.None
            };

            // Save the document as MHTML using the configured options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
