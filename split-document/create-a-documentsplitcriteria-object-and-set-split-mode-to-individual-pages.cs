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
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, true);
            Directory.CreateDirectory(outputDir);

            // Create a sample document with three pages.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Page 1 content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2 content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3 content.");

            // Create a DocumentSplitCriteria instance and set it to split at page breaks.
            DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.PageBreak;

            // Configure HtmlSaveOptions to use the split criteria.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = splitCriteria
            };

            // Save the document. Aspose.Words will create separate HTML files for each page.
            string baseFileName = Path.Combine(outputDir, "SplitDocument.html");
            doc.Save(baseFileName, saveOptions);

            // Verify that the expected number of split files were created.
            // The base file plus one file per page (three pages => three parts).
            string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument*")
                                            .Where(f => f.EndsWith(".html", StringComparison.OrdinalIgnoreCase))
                                            .OrderBy(f => f)
                                            .ToArray();

            // Expecting three split parts (the base file is the first part).
            if (splitFiles.Length != 3)
                throw new InvalidOperationException($"Expected 3 split files, but found {splitFiles.Length}.");

            // Output the names of the generated files (optional, for demonstration).
            foreach (string file in splitFiles)
                Console.WriteLine($"Generated: {Path.GetFileName(file)}");
        }
    }
}
