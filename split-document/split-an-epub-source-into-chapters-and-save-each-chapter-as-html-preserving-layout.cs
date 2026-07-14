using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitEpubIntoHtmlChapters
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document with heading paragraphs that will act as chapters.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the first chapter. It contains some sample text to demonstrate splitting.");

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the second chapter. More sample text follows.");

        // Chapter 3
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Third chapter content goes here.");

        // Save the document as EPUB – this will be our source file.
        string epubPath = Path.Combine(outputDir, "Sample.epub");
        sampleDoc.Save(epubPath, SaveFormat.Epub);

        // -----------------------------------------------------------------
        // 2. Load the EPUB and split it into separate HTML files per chapter.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(epubPath);

        // Configure HTML save options to split at heading paragraphs (level 1 = chapters).
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 1 // split only at Heading 1 styles.
        };

        // Base name for the first HTML file; additional parts will be created automatically.
        string baseHtmlPath = Path.Combine(outputDir, "Chapter.html");
        epubDoc.Save(baseHtmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 3. Validate that the expected split files were created.
        // -----------------------------------------------------------------
        // The main file plus one file per chapter (the first chapter is in the main file).
        string[] htmlFiles = Directory.GetFiles(outputDir, "Chapter*.html")
                                      .OrderBy(f => f)
                                      .ToArray();

        // We expect at least three files: Chapter.html, Chapter-01.html, Chapter-02.html, Chapter-03.html.
        // The exact naming depends on Aspose.Words version; we check that the count matches the number of chapters.
        int expectedChapterCount = 3; // we created three chapters.
        if (htmlFiles.Length < expectedChapterCount)
        {
            throw new InvalidOperationException(
                $"Expected at least {expectedChapterCount} HTML files after splitting, but found {htmlFiles.Length}.");
        }

        // Optional: output the list of generated files.
        Console.WriteLine("Generated HTML chapter files:");
        foreach (string file in htmlFiles)
        {
            Console.WriteLine(" - " + Path.GetFileName(file));
        }
    }
}
