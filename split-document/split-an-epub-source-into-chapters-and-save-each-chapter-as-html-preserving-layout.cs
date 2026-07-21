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
        // 1. Create a sample document that will be saved as EPUB.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // First chapter.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the first chapter.");

        // Second chapter.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the second chapter.");

        // Save the sample as EPUB – this will be our source file.
        string sourceEpubPath = Path.Combine(artifactsDir, "Source.epub");
        sampleDoc.Save(sourceEpubPath, SaveFormat.Epub);

        // -----------------------------------------------------------------
        // 2. Load the EPUB source document.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(sourceEpubPath);

        // -----------------------------------------------------------------
        // 3. Configure HtmlSaveOptions to split at heading paragraphs.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 1 // Split at Heading 1 only.
        };

        // Base name for the output HTML files.
        string baseHtmlName = Path.Combine(artifactsDir, "Chapter.html");

        // Save the document; Aspose.Words will create multiple HTML files.
        epubDoc.Save(baseHtmlName, htmlOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the expected split files were created.
        // -----------------------------------------------------------------
        // The first part keeps the original name, subsequent parts get a suffix like "-01.html".
        string[] splitFiles = Directory.GetFiles(artifactsDir, "Chapter*.html")
                                       .OrderBy(f => f)
                                       .ToArray();

        // Expect at least two files (two chapters).
        if (splitFiles.Length < 2)
        {
            throw new InvalidOperationException($"Expected at least 2 split HTML files, but found {splitFiles.Length}.");
        }

        // Output the list of generated files (optional, for verification).
        Console.WriteLine("Generated HTML chapter files:");
        foreach (string file in splitFiles)
        {
            Console.WriteLine($"- {Path.GetFileName(file)}");
        }
    }
}
