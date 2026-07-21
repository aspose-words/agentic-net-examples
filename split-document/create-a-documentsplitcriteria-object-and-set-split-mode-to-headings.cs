using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentByHeadings
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with heading paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading level 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        // Heading level 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        // Heading level 3 (will not be a split point if we limit to level 2)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");

        // Another heading level 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 4");

        // Heading level 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 5");

        // Heading level 3
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 6");

        // Create a DocumentSplitCriteria value that splits at heading paragraphs.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.HeadingParagraph;

        // Configure HtmlSaveOptions to use the split criteria.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = splitCriteria,
            DocumentSplitHeadingLevel = 2 // split at headings 1 and 2
        };

        // Save the document; Aspose.Words will generate multiple HTML files.
        string mainFileName = Path.Combine(outputDir, "SplitByHeadings.html");
        doc.Save(mainFileName, saveOptions);

        // Validate that the expected split parts were created.
        // The base file is "SplitByHeadings.html", subsequent parts are "-01.html", "-02.html", etc.
        string[] expectedParts = {
            mainFileName,
            Path.Combine(outputDir, "SplitByHeadings-01.html"),
            Path.Combine(outputDir, "SplitByHeadings-02.html"),
            Path.Combine(outputDir, "SplitByHeadings-03.html")
        };

        foreach (string partPath in expectedParts)
        {
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split part not found: {partPath}");
        }

        // Indicate successful completion (no console output required).
    }
}
