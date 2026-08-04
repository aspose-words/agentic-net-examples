using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define folders for input and output.
        string baseDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with three sections.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3");

        // Save the source document to a temporary file.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Define split criteria (split at each section break) and save.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Base file name for the split operation.
        string splitBaseName = Path.Combine(outputDir, "SplitDocument.html");
        loadedDoc.Save(splitBaseName, saveOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the split parts were created.
        // -----------------------------------------------------------------
        // Expected files: SplitDocument.html, SplitDocument-01.html, SplitDocument-02.html
        string[] expectedFiles =
        {
            splitBaseName,
            Path.Combine(outputDir, "SplitDocument-01.html"),
            Path.Combine(outputDir, "SplitDocument-02.html")
        };

        foreach (string filePath in expectedFiles)
        {
            if (!File.Exists(filePath))
                throw new Exception($"Expected split file not found: {filePath}");
        }

        // All split files exist – the workflow completed successfully.
        Console.WriteLine("Document split completed. Generated files:");
        foreach (string filePath in expectedFiles)
            Console.WriteLine($"  {Path.GetFileName(filePath)}");
    }
}
