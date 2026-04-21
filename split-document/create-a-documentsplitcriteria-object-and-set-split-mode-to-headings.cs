using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a sample document containing headings of various levels.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 4");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 5");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 6");

        // Create a DocumentSplitCriteria value that splits at heading paragraphs.
        DocumentSplitCriteria criteria = DocumentSplitCriteria.HeadingParagraph;

        // Configure HTML save options to use the split criteria and limit heading level to 2.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = criteria,
            DocumentSplitHeadingLevel = 2 // split at Heading 1 and Heading 2.
        };

        // Save the document; Aspose.Words will generate several HTML files.
        string baseFileName = Path.Combine(outputDir, "SplitByHeadings.html");
        doc.Save(baseFileName, saveOptions);

        // Expected split files (the first part keeps the original name,
        // subsequent parts receive a numeric suffix).
        string[] expectedFiles =
        {
            baseFileName,
            Path.Combine(outputDir, "SplitByHeadings-01.html"),
            Path.Combine(outputDir, "SplitByHeadings-02.html"),
            Path.Combine(outputDir, "SplitByHeadings-03.html")
        };

        // Verify that each expected file exists.
        foreach (string filePath in expectedFiles)
        {
            if (!File.Exists(filePath))
                throw new Exception($"Split output not found: {filePath}");
        }

        // Load one of the split parts and write its text to the console.
        Document part = new Document(expectedFiles[1]);
        Console.WriteLine("Content of first split part:");
        Console.WriteLine(part.GetText().Trim());
    }
}
