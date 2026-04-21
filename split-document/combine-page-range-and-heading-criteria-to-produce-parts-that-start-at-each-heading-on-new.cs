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

        // Build a sample document with headings and explicit page breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading (Heading 1) on a new page.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");
        builder.InsertBreak(BreakType.PageBreak);

        // Second heading (Heading 2) on a new page.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of section 1.1.");
        builder.InsertBreak(BreakType.PageBreak);

        // Third heading (Heading 2) on a new page.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of section 1.2.");

        // Configure split options: split at headings (levels 1‑2) and at page breaks.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.PageBreak,
            DocumentSplitHeadingLevel = 2 // split at Heading 1 and Heading 2.
        };

        // Save the document; Aspose.Words will generate separate HTML files for each part.
        string baseFileName = Path.Combine(outputDir, "CombinedSplit.html");
        doc.Save(baseFileName, saveOptions);

        // Validate that the base file exists.
        ValidateFileExists(baseFileName);

        // Validate and list the generated split parts (e.g., CombinedSplit-01.html, CombinedSplit-02.html, …).
        for (int i = 1; i <= 10; i++)
        {
            string partFile = Path.Combine(outputDir, $"CombinedSplit-{i:D2}.html");
            if (File.Exists(partFile))
            {
                Console.WriteLine($"Created split part: {Path.GetFileName(partFile)}");
            }
        }
    }

    private static void ValidateFileExists(string path)
    {
        if (!File.Exists(path))
            throw new Exception($"Expected file not found: {path}");
    }
}
