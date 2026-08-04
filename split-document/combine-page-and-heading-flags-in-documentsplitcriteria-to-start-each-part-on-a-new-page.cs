using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define an output folder relative to the current directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Base file name for the HTML output.
        string baseFileName = Path.Combine(outputDir, "SplitDocument.html");

        // Create a sample document with headings and explicit page breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading (level 1) and some content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 1.");

        // Insert a page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Second heading (level 2) and some content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 2.");

        // Insert another page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Third heading (level 3) and some content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 3.");

        // Configure HTML save options to split on both page breaks and heading paragraphs.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2 // Split at heading levels 1 and 2.
        };

        // Save the document; Aspose.Words will create multiple HTML files.
        doc.Save(baseFileName, options);

        // Verify that multiple split files were created.
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument*.html")
                                       .OrderBy(f => f)
                                       .ToArray();

        if (splitFiles.Length < 2)
            throw new Exception("Expected multiple split HTML files, but only one was found.");

        // Output the list of generated files.
        foreach (string file in splitFiles)
            Console.WriteLine("Created: " + file);
    }
}
