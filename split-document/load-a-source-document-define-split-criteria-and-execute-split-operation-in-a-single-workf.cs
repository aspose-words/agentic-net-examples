using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    public class Program
    {
        public static void Main()
        {
            // Define an output folder for all generated files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a sample source document with headings and regular text.
            // -----------------------------------------------------------------
            string sourcePath = Path.Combine(outputDir, "Source.docx");
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            // Heading level 1
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading 1");

            // Heading level 2
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Heading 2");

            // Heading level 3 (will not be a split point because we split only up to level 2)
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("Heading 3");

            // Another heading level 1
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading 4");

            // Another heading level 2
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Heading 5");

            // Save the source document to disk.
            sourceDoc.Save(sourcePath);

            // -----------------------------------------------------------------
            // 2. Load the source document.
            // -----------------------------------------------------------------
            Document doc = new Document(sourcePath);

            // -----------------------------------------------------------------
            // 3. Define split criteria – split at heading paragraphs up to level 2.
            // -----------------------------------------------------------------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                DocumentSplitHeadingLevel = 2 // Split at Heading 1 and Heading 2.
            };

            // Base file name for the split output. Aspose.Words will create additional
            // files with suffixes like "-01.html", "-02.html", etc.
            string baseFileName = Path.Combine(outputDir, "SplitDocument.html");

            // -----------------------------------------------------------------
            // 4. Save the document using the split options.
            // -----------------------------------------------------------------
            doc.Save(baseFileName, saveOptions);

            // -----------------------------------------------------------------
            // 5. Validate that the split operation produced multiple files.
            // -----------------------------------------------------------------
            // All files start with "SplitDocument" and have the .html extension.
            string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument*.html");

            // Expect at least the base file and one additional part.
            if (splitFiles.Length < 2)
            {
                throw new InvalidOperationException(
                    $"Expected multiple split files, but only found {splitFiles.Length} file(s).");
            }

            // Optional: output the list of generated files for verification.
            Console.WriteLine("Split operation completed. Generated files:");
            foreach (string file in splitFiles.OrderBy(f => f))
            {
                Console.WriteLine($" - {Path.GetFileName(file)}");
            }
        }
    }
}
