using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentByHeading
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading level 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        // Heading level 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of section 1.1.");

        // Another heading level 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Heading level 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of section 2.1.");

        // Create HtmlSaveOptions and set split criteria to headings.
        DocumentSplitCriteria criteria = DocumentSplitCriteria.HeadingParagraph;
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = criteria,
            DocumentSplitHeadingLevel = 2 // split at Heading 1 and Heading 2.
        };

        // Save the document; this will produce multiple HTML files.
        string baseFileName = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(baseFileName, saveOptions);

        // Validate that split files were created.
        // Expected files: SplitDocument.html, SplitDocument-01.html, SplitDocument-02.html, etc.
        string[] expectedFiles = Directory.GetFiles(outputDir, "SplitDocument*.html");
        if (expectedFiles.Length < 2)
        {
            throw new InvalidOperationException("Expected multiple split HTML files, but fewer were found.");
        }

        // Optional: list the generated files.
        foreach (string file in expectedFiles)
        {
            Console.WriteLine($"Generated: {Path.GetFileName(file)}");
        }
    }
}
