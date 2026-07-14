using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with several heading levels.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

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

        // Save the source document (optional, demonstrates loading from file).
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document.
        // -----------------------------------------------------------------
        Document doc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Define split criteria – split at heading paragraphs up to level 2.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2
        };

        // -----------------------------------------------------------------
        // 4. Execute the split operation by saving with the defined options.
        // -----------------------------------------------------------------
        string outputBase = Path.Combine(artifactsDir, "SplitDocument.html");
        doc.Save(outputBase, saveOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the expected split files were created.
        // -----------------------------------------------------------------
        string[] expectedFiles =
        {
            outputBase,
            Path.Combine(artifactsDir, "SplitDocument-01.html"),
            Path.Combine(artifactsDir, "SplitDocument-02.html"),
            Path.Combine(artifactsDir, "SplitDocument-03.html")
        };

        foreach (string file in expectedFiles)
        {
            if (!File.Exists(file))
                throw new FileNotFoundException($"Expected split file not found: {file}");
        }

        // Indicate successful completion.
        Console.WriteLine("Document split completed successfully.");
    }
}
