using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Create a sample document with two chapters separated by a page break.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Chapter 1 (Heading 1)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        // Insert a manual page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 2 (Heading 1)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Ensure the output folder exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Configure HTML save options:
        // - Split on both page breaks and heading paragraphs.
        // - Use heading level 1 as the split point.
        // - Do not export headers/footers (they are not needed for this demo).
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 1,
            ExportHeadersFootersMode = ExportHeadersFootersMode.None
        };

        // Save the document; Aspose.Words will generate multiple HTML files.
        string baseFilePath = Path.Combine(outputDir, "SampleDocument.html");
        doc.Save(baseFilePath, saveOptions);

        // Validate that at least one split file (the base file) was created.
        // Additional split parts, if any, will have an index suffix (e.g., SampleDocument_1.html).
        string[] allHtmlFiles = Directory.GetFiles(outputDir, "SampleDocument*.html");
        if (allHtmlFiles.Length == 0)
        {
            throw new InvalidOperationException("No HTML files were generated.");
        }

        // Optionally, list the generated files (uncomment for debugging).
        // foreach (var file in allHtmlFiles) Console.WriteLine(file);
    }
}
