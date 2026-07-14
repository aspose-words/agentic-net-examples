using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading (level 1) and some text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 1.");

        // Insert an explicit page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Second heading (level 2) and some text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 2.");

        // Another page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Third heading (level 1) and some text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 3.");

        // Configure HtmlSaveOptions to split on both page breaks and heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2 // Split at headings up to level 2.
        };

        // Save the document. This will produce the main file and additional split parts.
        string baseFileName = "CombinedSplit.html";
        string baseFilePath = Path.Combine(outputDir, baseFileName);
        doc.Save(baseFilePath, saveOptions);

        // Validate that split parts were created.
        string baseNameWithoutExt = Path.GetFileNameWithoutExtension(baseFileName);
        string[] splitFiles = Directory.GetFiles(outputDir, $"{baseNameWithoutExt}*.html")
                                       .OrderBy(f => f)
                                       .ToArray();

        // Expect at least the main file and one split part.
        if (splitFiles.Length < 2)
            throw new InvalidOperationException("Expected split output files were not created.");

        // Output the list of generated files.
        Console.WriteLine("Generated HTML parts:");
        foreach (string file in splitFiles)
            Console.WriteLine(Path.GetFileName(file));
    }
}
