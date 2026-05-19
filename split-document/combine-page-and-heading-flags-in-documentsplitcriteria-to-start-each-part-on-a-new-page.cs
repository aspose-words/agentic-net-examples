using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with headings and explicit page breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading and some content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 1.");

        // Explicit page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Second heading and content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 2.");

        // Another page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Third heading and content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 3.");

        // Configure HtmlSaveOptions to split the document at page breaks AND heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph,
            // Split at all heading levels (1‑9) to ensure headings are split points.
            DocumentSplitHeadingLevel = 9
        };

        // Save the document. The main file will be "SplitDocument.html",
        // additional parts will be named "SplitDocument-01.html", "SplitDocument-02.html", etc.
        string baseFileName = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(baseFileName, saveOptions);

        // Verify that the split parts were created.
        // The base file plus any "-NN.html" files constitute the split output.
        List<string> splitFiles = new List<string> { baseFileName };
        splitFiles.AddRange(Directory.GetFiles(outputDir, "SplitDocument-*.html"));

        if (splitFiles.Count < 2)
        {
            throw new Exception("Expected multiple split files, but only one was found.");
        }

        Console.WriteLine($"Created {splitFiles.Count} split HTML files:");
        foreach (string filePath in splitFiles)
        {
            Console.WriteLine(filePath);
        }
    }
}
