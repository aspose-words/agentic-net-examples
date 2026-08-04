using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        // Create a DocumentSplitCriteria instance and set it to split at page breaks,
        // which effectively creates a separate part for each page when saved to HTML.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.PageBreak;

        // Configure HTML save options to use the split criteria.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = splitCriteria
        };

        // Save the document; Aspose.Words will generate separate HTML files for each page.
        string baseFileName = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(baseFileName, saveOptions);

        // Verify that multiple split files were created.
        // The main file plus additional parts will share the same base name.
        var splitFiles = Directory.GetFiles(outputDir, "SplitDocument*")
                                  .Where(f => f.EndsWith(".html", StringComparison.OrdinalIgnoreCase))
                                  .ToArray();

        if (splitFiles.Length < 2)
        {
            throw new InvalidOperationException("Expected multiple split HTML files, but only one was found.");
        }

        // Output the names of the generated files (optional, for demonstration).
        foreach (var file in splitFiles)
        {
            Console.WriteLine("Generated file: " + Path.GetFileName(file));
        }
    }
}
