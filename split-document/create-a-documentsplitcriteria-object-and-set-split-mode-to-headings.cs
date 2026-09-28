using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Manual splitting by Heading1 paragraphs.
        var paragraphs = doc.FirstSection.Body.Paragraphs;
        int chapterIndex = 0;

        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = paragraphs[i];
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                chapterIndex++;

                // Create a new document for this chapter.
                Document splitDoc = new Document();
                // Ensure the body is empty before adding content.
                splitDoc.FirstSection.Body.RemoveAllChildren();

                // Import the heading paragraph.
                Node importedHeading = splitDoc.ImportNode(para, true, ImportFormatMode.KeepSourceFormatting);
                splitDoc.FirstSection.Body.AppendChild(importedHeading);

                // Import following paragraphs until the next Heading1 or end of document.
                int j = i + 1;
                while (j < paragraphs.Count &&
                       paragraphs[j].ParagraphFormat.StyleIdentifier != StyleIdentifier.Heading1)
                {
                    Node importedPara = splitDoc.ImportNode(paragraphs[j], true, ImportFormatMode.KeepSourceFormatting);
                    splitDoc.FirstSection.Body.AppendChild(importedPara);
                    j++;
                }

                // Save the split chapter as HTML.
                string chapterFile = Path.Combine(outputDir, $"Sample_Chapter{chapterIndex}.html");
                splitDoc.Save(chapterFile, SaveFormat.Html);

                // Move the outer loop index to the last processed paragraph.
                i = j - 1;
            }
        }

        // Verify that split files were created.
        string[] htmlFiles = Directory.GetFiles(outputDir, "*.html");
        if (htmlFiles.Length < 2)
        {
            throw new InvalidOperationException(
                $"Expected multiple HTML files after splitting, but only {htmlFiles.Length} were found.");
        }

        Console.WriteLine($"Splitting completed. {htmlFiles.Length} HTML files created in '{outputDir}'.");
    }
}
