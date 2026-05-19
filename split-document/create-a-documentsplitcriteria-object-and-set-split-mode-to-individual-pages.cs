using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define folders for input and output.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content.");

        // Create a DocumentSplitCriteria value that splits at page breaks,
        // which effectively creates a separate file for each page in this example.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.PageBreak;

        // Configure HTML save options to use the split criteria.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = splitCriteria
        };

        // Save the document. The main file will be "SplitDocument.html"
        // and additional parts will be named "SplitDocument-01.html", etc.
        string mainFilePath = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(mainFilePath, saveOptions);

        // Validate that the expected split files were created.
        // The main file plus two additional parts (for three pages total).
        string[] expectedFiles = Directory.GetFiles(outputDir, "SplitDocument*")
                                          .Select(Path.GetFileName)
                                          .OrderBy(name => name)
                                          .ToArray();

        // Expect at least three files: the main file and two split parts.
        if (expectedFiles.Length < 3)
        {
            throw new InvalidOperationException(
                $"Expected at least 3 split files, but found {expectedFiles.Length}.");
        }

        // Output the names of the generated files (for demonstration purposes).
        Console.WriteLine("Generated split files:");
        foreach (string fileName in expectedFiles)
        {
            Console.WriteLine(fileName);
        }
    }
}
