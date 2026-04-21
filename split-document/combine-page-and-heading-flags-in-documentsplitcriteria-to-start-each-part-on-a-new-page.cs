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

        // Create a sample document with headings and explicit page breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading (level 1) – will be a split point.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");
        builder.Writeln("Content under heading 1.");

        // Insert an explicit page break – also a split point.
        builder.InsertBreak(BreakType.PageBreak);

        // Second heading (level 2) – will be a split point.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");
        builder.Writeln("Content under heading 2.");

        // Another page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Third heading (level 3) – will be a split point if the heading level is high enough.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");
        builder.Writeln("Content under heading 3.");

        // Save the document to HTML, splitting on both page breaks and heading paragraphs.
        string baseFileName = Path.Combine(outputDir, "SplitDocument.html");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 3 // Include headings up to level 3.
        };

        doc.Save(baseFileName, saveOptions);

        // Validate that split parts were created.
        // Aspose.Words creates files named: baseName-01.html, baseName-02.html, etc.
        string baseNameWithoutExt = Path.GetFileNameWithoutExtension(baseFileName);
        string[] expectedParts = { "-01.html", "-02.html", "-03.html", "-04.html" };
        foreach (string suffix in expectedParts)
        {
            string partPath = Path.Combine(outputDir, baseNameWithoutExt + suffix);
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split part not found: {partPath}");
        }

        // Indicate successful execution (no console output required by the task).
    }
}
