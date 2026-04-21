using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with heading paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading level 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        // Heading level 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of section 1.1.");

        // Heading level 1 again.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Configure HtmlSaveOptions to split by heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2 // Split at Heading 1 and Heading 2.
        };

        // Save the document; this will produce multiple HTML files.
        string baseFilePath = Path.Combine(outputDir, "SplitByHeading.html");
        doc.Save(baseFilePath, saveOptions);

        // Validate that the split files were created.
        // The main file and at least one split part should exist.
        if (!File.Exists(baseFilePath))
            throw new FileNotFoundException("Base HTML file was not created.", baseFilePath);

        // Split parts are named with a suffix "-01.html", "-02.html", etc.
        string part1Path = Path.Combine(outputDir, "SplitByHeading-01.html");
        if (!File.Exists(part1Path))
            throw new FileNotFoundException("First split HTML part was not created.", part1Path);

        // Optional: output confirmation (no interactive prompts required).
        Console.WriteLine("Document split successfully. Files created in: " + outputDir);
    }
}
