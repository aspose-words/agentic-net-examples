using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace CombinePageAndHeadingFlags
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set up input and output directories relative to the executable location.
            string baseDir = AppContext.BaseDirectory;
            string dataDir = Path.Combine(baseDir, "Data");
            string outputDir = Path.Combine(baseDir, "Output");

            Directory.CreateDirectory(dataDir);
            Directory.CreateDirectory(outputDir);

            // Path to the input document.
            string inputPath = Path.Combine(dataDir, "InputDocument.docx");

            // If the input file does not exist, create a simple document for demonstration.
            if (!File.Exists(inputPath))
            {
                Document tempDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(tempDoc);

                // First part – heading level 1.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                builder.Writeln("Chapter 1");

                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                builder.Writeln("This is the first chapter.");

                // Insert a page break to start a new part.
                builder.InsertBreak(BreakType.PageBreak);

                // Second part – heading level 2.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
                builder.Writeln("Section 1.1");

                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                builder.Writeln("Details of section 1.1.");

                // Save the temporary document.
                tempDoc.Save(inputPath);
            }

            // Load the document.
            Document doc = new Document(inputPath);

            // Configure HTML save options to split the output into separate parts.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph,
                DocumentSplitHeadingLevel = 2,
                // Ensure each part is saved to a separate file.
                ExportHeadersFootersMode = ExportHeadersFootersMode.None,
                ExportPageMargins = false
            };

            // Save the split document.
            string outputPath = Path.Combine(outputDir, "OutputDocument.html");
            doc.Save(outputPath, saveOptions);

            Console.WriteLine($"Document split and saved to: {outputPath}");
        }
    }
}
