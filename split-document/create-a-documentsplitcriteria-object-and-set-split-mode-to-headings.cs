using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing two heading paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First chapter.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        // Second chapter.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Configure HTML save options to split the document at heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Use the correct enum value for heading‑based splitting.
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
        // Optional: specify the heading level to split on (1 = Heading1).
        saveOptions.DocumentSplitHeadingLevel = 1;

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Save the document; Aspose.Words will generate separate HTML files for each heading.
        string baseFileName = Path.Combine(outputDir, "SplitDocument.html");
        doc.Save(baseFileName, saveOptions);

        // Verify that at least two split files were created.
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument*.html");
        if (splitFiles.Length < 2)
        {
            throw new InvalidOperationException(
                $"Expected multiple split HTML files, but only {splitFiles.Length} were found.");
        }

        // List the generated files.
        foreach (string file in splitFiles)
        {
            Console.WriteLine($"Generated: {file}");
        }
    }
}
