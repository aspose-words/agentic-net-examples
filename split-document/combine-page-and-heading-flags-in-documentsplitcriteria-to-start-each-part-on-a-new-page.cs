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

        // Build a sample document containing headings and explicit page breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        // Normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the first chapter.");

        // First page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Heading 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Details of section 1.1.");

        // Second page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Another Heading 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the second chapter.");

        // Configure HtmlSaveOptions to split on both page breaks and heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2 // Split at Heading 1 and Heading 2.
        };

        // Save the document; Aspose.Words will generate several HTML files.
        string mainFilePath = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(mainFilePath, saveOptions);

        // Verify that split parts were created.
        var splitFiles = Directory.GetFiles(outputDir, "SplitDocument-*.html");
        if (!splitFiles.Any())
        {
            throw new InvalidOperationException("No split HTML files were generated.");
        }

        // List generated files (optional).
        Console.WriteLine("Generated files:");
        Console.WriteLine(mainFilePath);
        foreach (string file in splitFiles)
        {
            Console.WriteLine(file);
        }
    }
}
