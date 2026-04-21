using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with several heading levels.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading #1");

        // Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading #2");

        // Heading 3 (will not be a split point because we limit to level 2)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading #3");

        // Another Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading #4");

        // Another Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading #5");

        // Another Heading 3
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading #6");

        // -----------------------------------------------------------------
        // 2. Configure split options – split at heading paragraphs up to level 2.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2,
            SaveFormat = SaveFormat.Html
        };

        // -----------------------------------------------------------------
        // 3. Save the document; Aspose.Words will create multiple HTML files.
        // -----------------------------------------------------------------
        string mainFilePath = Path.Combine(outputDir, "SplitDocument.html");
        sourceDoc.Save(mainFilePath, saveOptions);

        // -----------------------------------------------------------------
        // 4. Verify that the expected split files were created.
        // -----------------------------------------------------------------
        string[] expectedFiles =
        {
            mainFilePath,
            Path.Combine(outputDir, "SplitDocument-01.html"),
            Path.Combine(outputDir, "SplitDocument-02.html"),
            Path.Combine(outputDir, "SplitDocument-03.html")
        };

        foreach (string file in expectedFiles)
        {
            if (!File.Exists(file))
                throw new FileNotFoundException($"Expected split file was not created: {file}");
        }

        // All split files exist – the workflow completed successfully.
    }
}
