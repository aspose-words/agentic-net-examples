using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Create a DocumentSplitCriteria instance and set it to split at page breaks.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.PageBreak;

        // Configure HtmlSaveOptions to use the split criteria.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = splitCriteria
        };

        // Save the document; Aspose.Words will generate separate HTML files for each page.
        string baseFileName = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(baseFileName, saveOptions);

        // Verify that split files were created (the base file plus additional parts).
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument*.html");
        if (splitFiles.Length < 2)
            throw new Exception("Split HTML files were not created as expected.");

        Console.WriteLine($"Successfully created {splitFiles.Length} split HTML files in '{outputDir}'.");
    }
}
