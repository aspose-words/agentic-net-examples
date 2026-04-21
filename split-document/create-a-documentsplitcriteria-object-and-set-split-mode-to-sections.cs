using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains three separate sections.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage); // Start Section 2
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage); // Start Section 3
        builder.Writeln("Content of Section 3");

        // -----------------------------------------------------------------
        // 2. Create a DocumentSplitCriteria object and set it to split by sections.
        // -----------------------------------------------------------------
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.SectionBreak;

        // -----------------------------------------------------------------
        // 3. Configure HtmlSaveOptions to use the split criteria.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = splitCriteria
        };

        // -----------------------------------------------------------------
        // 4. Save the document. Aspose.Words will generate multiple HTML files,
        //    one for each section.
        // -----------------------------------------------------------------
        string baseFileName = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(baseFileName, saveOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the split output files were created.
        // -----------------------------------------------------------------
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument*.html");
        // Expect at least two files: the base file and one additional part.
        if (splitFiles.Length < 2)
        {
            throw new InvalidOperationException(
                $"Expected multiple split files, but only {splitFiles.Length} were found.");
        }

        // Optional: list the generated files (for debugging purposes).
        Console.WriteLine("Generated split files:");
        foreach (string file in splitFiles.OrderBy(f => f))
        {
            Console.WriteLine($" - {Path.GetFileName(file)}");
        }
    }
}
