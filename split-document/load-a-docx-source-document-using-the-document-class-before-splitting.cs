using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define paths for the sample source document and the split output.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);
        string sourcePath = Path.Combine(dataDir, "SourceDocument.docx");
        string outputBasePath = Path.Combine(dataDir, "SplitOutput.html");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document with multiple sections.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section content.
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section content.
        builder.Writeln("Section 2 - Paragraph 1");
        builder.Writeln("Section 2 - Paragraph 2");

        // Save the sample document.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX source document using the Document class.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Split the loaded document by sections using HtmlSaveOptions.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Saving with the split criteria will generate multiple HTML files:
        // "SplitOutput.html", "SplitOutput-01.html", etc.
        loadedDoc.Save(outputBasePath, saveOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the split output files were created.
        // -----------------------------------------------------------------
        string outputDirectory = Path.GetDirectoryName(outputBasePath);
        string outputFileNameWithoutExt = Path.GetFileNameWithoutExtension(outputBasePath);
        string[] splitFiles = Directory.GetFiles(outputDirectory, $"{outputFileNameWithoutExt}*.html");

        // Expect at least two files (original + one split part).
        if (splitFiles.Length < 2)
        {
            throw new InvalidOperationException("Document splitting failed: expected multiple output files.");
        }

        // Optional: display the generated file names (not required for non‑interactive run).
        foreach (string file in splitFiles)
        {
            Console.WriteLine($"Generated: {Path.GetFileName(file)}");
        }
    }
}
