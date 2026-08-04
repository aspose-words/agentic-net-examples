using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document with heading paragraphs representing chapters.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Save the document as an EPUB file (the source to be split).
        string epubPath = Path.Combine(artifactsDir, "Sample.epub");
        sampleDoc.Save(epubPath, SaveFormat.Epub);

        // -----------------------------------------------------------------
        // 2. Load the EPUB source.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(epubPath);

        // -----------------------------------------------------------------
        // 3. Configure HtmlSaveOptions to split at heading paragraphs.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 1 // Split only at Heading 1 (chapters).
        };

        // Optional: specify a folder for images extracted during HTML conversion.
        string imagesDir = Path.Combine(artifactsDir, "Images");
        if (Directory.Exists(imagesDir))
            Directory.Delete(imagesDir, true);
        Directory.CreateDirectory(imagesDir);
        saveOptions.ImagesFolder = imagesDir;

        // -----------------------------------------------------------------
        // 4. Save the EPUB as HTML with splitting enabled.
        // -----------------------------------------------------------------
        string htmlBasePath = Path.Combine(artifactsDir, "SplitOutput.html");
        epubDoc.Save(htmlBasePath, saveOptions);

        // -----------------------------------------------------------------
        // 5. Validate that split HTML files were created.
        // -----------------------------------------------------------------
        string[] htmlFiles = Directory.GetFiles(artifactsDir, "SplitOutput*.html");
        if (htmlFiles.Length < 2)
            throw new InvalidOperationException("Expected multiple HTML parts after splitting, but fewer were found.");

        // (Optional) Output the list of generated files for verification.
        foreach (string file in htmlFiles.OrderBy(f => f))
        {
            Console.WriteLine($"Generated: {Path.GetFileName(file)}");
        }
    }
}
