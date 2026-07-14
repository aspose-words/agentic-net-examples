using System;
using System.IO;
using Aspose.Words; // Core Aspose.Words namespace (contains Document, DocumentBuilder, NodeImporter, etc.)

public class Program
{
    public static void Main()
    {
        // Define an output directory for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with three sections.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("Section 1 - First paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - First paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - First paragraph.");

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document using the Document class.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Split the loaded document by its sections.
        //    Each section is saved as an independent DOCX file.
        // -----------------------------------------------------------------
        for (int i = 0; i < loadedDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document splitDoc = new Document();

            // Remove the default empty section that a new Document contains.
            splitDoc.RemoveAllChildren();

            // Import the current section from the loaded document into the new document.
            // NodeImporter handles style, list and other metadata translation.
            NodeImporter importer = new NodeImporter(loadedDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedSection = importer.ImportNode(loadedDoc.Sections[i], true);

            // Append the imported section to the split document.
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = Path.Combine(outputDir, $"Section_{i + 1}.docx");
            splitDoc.Save(splitPath);

            // Simple validation that the file was created.
            if (!File.Exists(splitPath))
                throw new InvalidOperationException($"Failed to create split file: {splitPath}");
        }

        // -----------------------------------------------------------------
        // 4. Report the results.
        // -----------------------------------------------------------------
        Console.WriteLine($"Source document created at: {sourcePath}");
        Console.WriteLine("Split documents:");
        foreach (string file in Directory.GetFiles(outputDir, "Section_*.docx"))
        {
            Console.WriteLine($" - {file}");
        }
    }
}
