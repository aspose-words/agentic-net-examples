using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PreserveFieldsExample
{
    static void Main()
    {
        // Create a source document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add some content with an embedded field.
        builder.Writeln("This is a sample document.");
        builder.InsertField("MERGEFIELD Name \\* MERGEFORMAT");
        builder.Writeln();
        builder.Writeln("Page 1 content.");

        // Insert a page break to ensure there is a second page.
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");

        // Extract pages 1‑2 (zero‑based indices 0 and 1) from the source document.
        Document extractedDoc = sourceDoc.ExtractPages(0, 1);

        // Configure save options for DOCX output.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            // Preserve the current field results; do not recalculate them on save.
            UpdateFields = false,
            // Omit the Aspose.Words generator name from the saved file.
            ExportGeneratorName = false
        };

        // Save the extracted document to the current directory.
        string outputPath = "Extracted.docx";
        extractedDoc.Save(outputPath, saveOptions);

        Console.WriteLine($"Extracted document saved to: {outputPath}");
    }
}
