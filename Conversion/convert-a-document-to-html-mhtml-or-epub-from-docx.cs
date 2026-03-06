using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourceDocx = @"C:\Docs\SampleDocument.docx";

            // Ensure the source file exists.
            if (!File.Exists(sourceDocx))
                throw new FileNotFoundException("Source DOCX not found.", sourceDocx);

            // Load the DOCX document.
            Document doc = new Document(sourceDocx);

            // -----------------------------------------------------------------
            // 1. Convert to HTML
            // -----------------------------------------------------------------
            string htmlOutput = @"C:\Docs\SampleDocument.html";
            // Save using the overload that specifies the format directly.
            doc.Save(htmlOutput, SaveFormat.Html);

            // -----------------------------------------------------------------
            // 2. Convert to MHTML (Web archive)
            // -----------------------------------------------------------------
            string mhtmlOutput = @"C:\Docs\SampleDocument.mhtml";
            // Save using the overload that specifies the format directly.
            doc.Save(mhtmlOutput, SaveFormat.Mhtml);

            // -----------------------------------------------------------------
            // 3. Convert to EPUB
            // -----------------------------------------------------------------
            string epubOutput = @"C:\Docs\SampleDocument.epub";

            // For EPUB we can use HtmlSaveOptions to control additional settings.
            HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                // Example: export document properties into the EPUB package.
                ExportDocumentProperties = true,

                // Example: split the EPUB into separate HTML parts at each heading paragraph.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                // Example: set the encoding to UTF-8 without BOM.
                Encoding = new System.Text.UTF8Encoding(false)
            };

            // Save the document as EPUB using the configured options.
            doc.Save(epubOutput, epubOptions);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
