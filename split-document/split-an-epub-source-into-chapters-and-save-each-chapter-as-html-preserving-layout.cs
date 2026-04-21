using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document with two chapters (Heading 1 style).
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of chapter 1.");

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of chapter 2.");

        // Save the document as an EPUB file – this will be the source to split.
        string epubPath = Path.Combine(outputDir, "source.epub");
        sampleDoc.Save(epubPath, SaveFormat.Epub);

        // -----------------------------------------------------------------
        // 2. Load the EPUB source.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(epubPath);

        // -----------------------------------------------------------------
        // 3. Configure HTML save options to split at heading paragraphs.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Split the document into separate HTML files at each Heading 1.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            // Optional: limit split to heading level 1 only.
            DocumentSplitHeadingLevel = 1
        };

        // Save the EPUB as HTML with splitting. The base file name is "chapter.html".
        string htmlBasePath = Path.Combine(outputDir, "chapter.html");
        epubDoc.Save(htmlBasePath, htmlOptions);

        // -----------------------------------------------------------------
        // 4. Verify that split HTML files were created.
        // -----------------------------------------------------------------
        string[] htmlFiles = Directory.GetFiles(outputDir, "chapter*.html");
        if (htmlFiles.Length < 2)
        {
            throw new InvalidOperationException("Expected multiple HTML files after splitting, but only one was found.");
        }

        // (Optional) Output the list of generated files for debugging purposes.
        foreach (string file in htmlFiles)
        {
            Console.WriteLine("Generated: " + Path.GetFileName(file));
        }
    }
}
