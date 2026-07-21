using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document with headings and sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading (level 1).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        // Add some normal text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 1.");

        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second heading (level 2) in the new section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        // Add more text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 2.");

        // Insert another heading (level 3) – this will not trigger a split because we limit to level 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 3.");

        // Configure HtmlSaveOptions to split by both headings and sections.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.SectionBreak,
            DocumentSplitHeadingLevel = 2 // Split at Heading 1 and Heading 2.
        };

        // Save the document; Aspose.Words will create multiple HTML files.
        string mainFileName = Path.Combine(artifactsDir, "SplitDocument.html");
        doc.Save(mainFileName, saveOptions);

        // Verify that split parts were created.
        string[] splitFiles = Directory.GetFiles(artifactsDir, "SplitDocument-*.html");
        Console.WriteLine($"Main file: {Path.GetFileName(mainFileName)}");
        Console.WriteLine($"Number of split parts: {splitFiles.Length}");
        foreach (string file in splitFiles)
        {
            Console.WriteLine($" - {Path.GetFileName(file)}");
        }

        // Simple validation: ensure at least one split part exists.
        if (splitFiles.Length == 0)
        {
            throw new InvalidOperationException("No split parts were generated.");
        }
    }
}
