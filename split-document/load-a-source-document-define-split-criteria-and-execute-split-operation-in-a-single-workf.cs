using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample source document with headings.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        CreateSampleDocument(sourcePath);

        // Load the source document.
        Document doc = new Document(sourcePath);

        // Define split criteria: split at heading paragraphs up to level 2.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2
        };

        // Execute the split by saving the document.
        string mainOutput = Path.Combine(outputDir, "Split.html");
        doc.Save(mainOutput, options);

        // Validate that split parts were created.
        ValidateSplitOutputs(outputDir, "Split");
    }

    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add headings of various levels.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading #1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading #2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading #3");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading #4");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading #5");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading #6");

        doc.Save(filePath);
    }

    private static void ValidateSplitOutputs(string folder, string baseName)
    {
        // Expected files: Split.html, Split-01.html, Split-02.html, Split-03.html
        var files = Directory.GetFiles(folder, $"{baseName}*.html");
        if (files.Length < 4)
        {
            throw new InvalidOperationException($"Expected at least 4 HTML files after split, but found {files.Length}.");
        }

        Console.WriteLine("Split operation produced the following files:");
        foreach (var file in files.OrderBy(f => f))
        {
            Console.WriteLine(file);
        }
    }
}
