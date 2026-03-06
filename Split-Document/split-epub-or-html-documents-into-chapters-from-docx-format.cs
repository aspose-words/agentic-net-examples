using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsChapterSplit
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file that contains headings (Heading 1, Heading 2, etc.).
            string sourceDocxPath = @"C:\Docs\SourceDocument.docx";

            // Path where the split EPUB file will be saved.
            // The EPUB will contain separate HTML parts for each heading up to the specified level.
            string outputEpubPath = @"C:\Docs\SplitDocument.epub";

            // Load the DOCX document.
            Document doc = new Document(sourceDocxPath);

            // Configure save options for EPUB/HTML conversion.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                // Split the output at paragraphs that use heading styles.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                // Define up to which heading level the split should occur (e.g., Heading 1‑3).
                DocumentSplitHeadingLevel = 3,

                // Export built‑in and custom document properties into the EPUB package.
                ExportDocumentProperties = true,

                // Use UTF‑8 encoding for the generated HTML parts.
                Encoding = Encoding.UTF8
            };

            // Save the document as an EPUB file; the content will be split into chapters.
            doc.Save(outputEpubPath, saveOptions);
        }
    }
}
