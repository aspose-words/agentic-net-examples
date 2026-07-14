using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a sample document with headings that start on new pages.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 1; i <= 3; i++)
            {
                if (i > 1)
                {
                    // Ensure each heading begins on a new page.
                    builder.InsertBreak(BreakType.PageBreak);
                }

                // Insert a Heading 1 paragraph.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                builder.Writeln($"Heading {i}");

                // Insert some normal content under the heading.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                builder.Writeln($"Content for heading {i} - paragraph 1.");
                builder.Writeln($"Content for heading {i} - paragraph 2.");
            }

            // Save the source document (optional, for inspection).
            string sourcePath = Path.Combine(outputDir, "Source.docx");
            doc.Save(sourcePath);

            // Configure HtmlSaveOptions to split at page breaks and heading paragraphs.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph,
                DocumentSplitHeadingLevel = 1, // Split at Heading 1 level.
                ExportHeadersFootersMode = ExportHeadersFootersMode.None,
                DocumentPartSavingCallback = new PartSavingCallback(outputDir)
            };

            // Save the document; Aspose.Words will invoke the callback for each split part.
            string mainOutputPath = Path.Combine(outputDir, "Combined.html");
            doc.Save(mainOutputPath, saveOptions);

            // Validate that split parts were created.
            string[] partFiles = Directory.GetFiles(outputDir, "Part_*.html");
            if (partFiles.Length == 0)
                throw new Exception("No split parts were generated.");

            Console.WriteLine($"Document split into {partFiles.Length} parts:");
            foreach (string file in partFiles)
                Console.WriteLine($" - {Path.GetFileName(file)}");
        }

        // Callback that assigns deterministic filenames for each document part.
        private class PartSavingCallback : IDocumentPartSavingCallback
        {
            private readonly string _outputFolder;
            private int _partIndex = 0;

            public PartSavingCallback(string outputFolder)
            {
                _outputFolder = outputFolder;
            }

            void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
            {
                // Generate a simple sequential filename.
                string partFileName = $"Part_{++_partIndex}.html";

                // Set the filename (without path) and provide a stream for the part.
                args.DocumentPartFileName = partFileName;
                args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, partFileName), FileMode.Create);
                // The stream will be closed automatically after saving the part.
            }
        }
    }
}
