using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------------------
        // 1. Create a sample source document with three sections,
        //    each having a different page orientation.
        // -------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1 – portrait (default).
        builder.Writeln("Section 1 – Portrait");
        builder.Writeln("This is the first section.");

        // Section 2 – landscape.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("Section 2 – Landscape");
        builder.Writeln("This section is in landscape mode.");

        // Section 3 – portrait again.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.Writeln("Section 3 – Portrait");
        builder.Writeln("Back to portrait orientation.");

        // Save the source document for reference.
        string sourcePath = Path.Combine(outputDir, "SourceDocument.docx");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -------------------------------------------------------------
        // 2. Split the document by its sections.
        //    For each section we create a new Document, import the section,
        //    and save it as an independent DOCX file.
        // -------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // Remove the default empty section.

            // Clone the current section from the source document.
            Section clonedSection = sourceDoc.Sections[i].Clone();

            // Import the cloned section into the new document.
            Node importedSection = partDoc.ImportNode(clonedSection, true);
            partDoc.AppendChild(importedSection);

            // Build the output file name.
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");

            // Save the split part preserving all formatting and orientation.
            partDoc.Save(partPath, SaveFormat.Docx);

            // -------------------------------------------------------------
            // 3. Validation – ensure the file was created.
            // -------------------------------------------------------------
            if (!File.Exists(partPath))
                throw new Exception($"Failed to create split part: {partPath}");
        }

        // All split parts have been saved successfully.
        Console.WriteLine($"Document split into {sourceDoc.Sections.Count} parts. Files are located in: {outputDir}");
    }
}
