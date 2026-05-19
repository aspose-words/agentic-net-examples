using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentByHeadings
{
    public static void Main()
    {
        // Define output directory and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with heading paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        // Normal paragraph under heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        // Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of section 1.1.");

        // Heading 1 again
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Create a DocumentSplitCriteria value that splits at heading paragraphs.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.HeadingParagraph;

        // Configure HtmlSaveOptions to use the split criteria.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = splitCriteria,
            DocumentSplitHeadingLevel = 2 // Split at Heading 1 and Heading 2.
        };

        // Save the document; Aspose.Words will generate multiple HTML files.
        string mainFilePath = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(mainFilePath, saveOptions);

        // Verify that split parts were created.
        // The main file and additional parts have suffixes like "-01.html", "-02.html", etc.
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument*.html");
        if (splitFiles.Length < 2)
        {
            throw new InvalidOperationException("Expected multiple split HTML files, but only one was found.");
        }

        // Output the names of the generated files (optional, for debugging).
        foreach (string file in splitFiles)
        {
            Console.WriteLine("Generated file: " + Path.GetFileName(file));
        }
    }
}
