using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create an output folder for the generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Build a sample document that contains headings and an explicit page break.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        // Normal paragraph under Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 1.");

        // Explicit page break – this will also act as a split point.
        builder.InsertBreak(BreakType.PageBreak);

        // Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        // Normal paragraph under Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 2.");

        // -----------------------------------------------------------------
        // 2. Configure HtmlSaveOptions to split the document at both page breaks
        //    and heading paragraphs. The DocumentSplitHeadingLevel limits the heading
        //    levels that trigger a split (1‑2 in this case).
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2
        };

        // Save the document. Aspose.Words will create the main file and additional
        // parts (e.g., SplitDocument-01.html, SplitDocument-02.html, …).
        string baseFileName = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(baseFileName, saveOptions);

        // -----------------------------------------------------------------
        // 3. Validate that the expected split files exist.
        // -----------------------------------------------------------------
        string part0 = baseFileName;                                 // original file
        string part1 = Path.Combine(outputDir, "SplitDocument-01.html"); // after first split point
        string part2 = Path.Combine(outputDir, "SplitDocument-02.html"); // after second split point

        if (!File.Exists(part0) || !File.Exists(part1) || !File.Exists(part2))
            throw new Exception("One or more expected split HTML parts were not created.");

        // Output the locations of the generated files (optional, not required for automation).
        Console.WriteLine("Document successfully split into parts:");
        Console.WriteLine(part0);
        Console.WriteLine(part1);
        Console.WriteLine(part2);
    }
}
