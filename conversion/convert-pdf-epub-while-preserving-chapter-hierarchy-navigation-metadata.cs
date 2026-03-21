using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfToEpubConversion
{
    class Program
    {
        static void Main()
        {
            // Create a simple document in memory with heading paragraphs to demonstrate chapter splitting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Chapter 1
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is the content of the first chapter.");

            // Chapter 2
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is the content of the second chapter.");

            // Define output path (using the system's temporary folder).
            string epubPath = Path.Combine(Path.GetTempPath(), "ResultDocument.epub");

            // Configure save options for EPUB output.
            HtmlSaveOptions epubSaveOptions = new HtmlSaveOptions
            {
                SaveFormat = SaveFormat.Epub,
                Encoding = new UTF8Encoding(false),
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                ExportDocumentProperties = true,
                NavigationMapLevel = 3
            };

            // Save the document as an EPUB using the configured options.
            doc.Save(epubPath, epubSaveOptions);

            Console.WriteLine($"EPUB file saved to: {epubPath}");
        }
    }
}
