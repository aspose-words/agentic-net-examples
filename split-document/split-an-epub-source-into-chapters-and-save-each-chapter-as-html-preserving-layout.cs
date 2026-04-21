using System;
using System.IO;
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
        // 1. Create a sample document with three chapters (Heading 1 style).
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        for (int i = 1; i <= 3; i++)
        {
            // Insert a heading that will act as a chapter title.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // Insert some body text for the chapter.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}. It contains several sentences to illustrate the split operation.");
        }

        // Save the document as an EPUB file – this will be the source we split.
        string epubPath = Path.Combine(outputDir, "Sample.epub");
        sampleDoc.Save(epubPath, SaveFormat.Epub);

        // ---------------------------------------------------------------
        // 2. Load the EPUB file we just created.
        // ---------------------------------------------------------------
        Document epubDoc = new Document(epubPath);

        // ---------------------------------------------------------------
        // 3. Configure HtmlSaveOptions to split at Heading 1 paragraphs.
        // ---------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 1, // Split at Heading 1 only.
            ExportDocumentProperties = true   // Preserve document properties in the output.
        };

        // Base file name for the split HTML parts.
        string baseHtmlPath = Path.Combine(outputDir, "Chapter.html");

        // Save the EPUB as HTML; Aspose.Words will generate separate files for each chapter.
        epubDoc.Save(baseHtmlPath, htmlOptions);

        // ---------------------------------------------------------------
        // 4. Validate that the expected split files exist.
        // ---------------------------------------------------------------
        // The first part uses the base name, subsequent parts get a numeric suffix.
        string part0 = baseHtmlPath;                     // Chapter.html
        string part1 = Path.Combine(outputDir, "Chapter-01.html");
        string part2 = Path.Combine(outputDir, "Chapter-02.html");

        if (!File.Exists(part0) || !File.Exists(part1) || !File.Exists(part2))
            throw new InvalidOperationException("One or more expected HTML chapter files were not created.");

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("EPUB successfully split into HTML chapters:");
        Console.WriteLine(part0);
        Console.WriteLine(part1);
        Console.WriteLine(part2);
    }
}
