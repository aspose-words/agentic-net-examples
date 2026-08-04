using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Create an output folder for the split parts.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        Directory.CreateDirectory(outputDir);

        // Build a sample document that contains three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("Section 1 - First paragraph.");
        builder.Writeln("Section 1 - Second paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage); // start Section 2

        // Section 2
        builder.Writeln("Section 2 - First paragraph.");
        builder.Writeln("Section 2 - Second paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage); // start Section 3

        // Section 3
        builder.Writeln("Section 3 - First paragraph.");
        builder.Writeln("Section 3 - Second paragraph.");

        // Split the document by its sections.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            Section sourceSection = doc.Sections[i];

            // Create a new empty document that will hold the imported section.
            Document splitDoc = new Document();

            // Remove the default empty section that Aspose.Words creates for a new document.
            splitDoc.RemoveAllChildren();

            // Import the source section into the new document.
            // NodeImporter works with Document objects, not with Section objects directly.
            NodeImporter importer = new NodeImporter(doc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sourceSection, true);

            // Append the imported section to the split document.
            splitDoc.AppendChild(importedSection);

            // Save the split part to a file.
            string partPath = Path.Combine(outputDir, $"SplitPart_{i + 1}.docx");
            splitDoc.Save(partPath, SaveFormat.Docx);

            // Verify that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split file: {partPath}");
        }

        Console.WriteLine($"Document split into {doc.Sections.Count} parts. Files saved to: {outputDir}");
    }
}
