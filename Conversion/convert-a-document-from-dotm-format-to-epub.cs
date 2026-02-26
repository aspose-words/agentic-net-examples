using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class DotmToEpubConverter
{
    static void Main()
    {
        // Load the macro‑enabled template (DOTM) from disk.
        Document doc = new Document("input.dotm");

        // Configure save options for EPUB output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Specify that the target format is EPUB.
            SaveFormat = SaveFormat.Epub,

            // Use UTF‑8 encoding (without BOM) for the generated EPUB.
            Encoding = Encoding.UTF8,

            // Export built‑in and custom document properties into the EPUB.
            ExportDocumentProperties = true,

            // Optional: split the EPUB into separate HTML parts at heading paragraphs.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph
        };

        // Save the document as an EPUB file using the configured options.
        doc.Save("output.epub", saveOptions);
    }
}
