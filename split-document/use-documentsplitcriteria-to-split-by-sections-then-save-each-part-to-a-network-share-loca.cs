using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the sample source document.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeSplitSample");
        Directory.CreateDirectory(tempFolder);

        // Path for the source document.
        string sourcePath = Path.Combine(tempFolder, "SourceDocument.docx");

        // Build a sample document with three sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("Section 1 - Content");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - Content");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - Content");

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // Load the document (simulating a real load scenario).
        Document loadedDoc = new Document(sourcePath);

        // Define a network share path (simulated with a local folder).
        string networkSharePath = Path.Combine(Path.GetTempPath(), "NetworkShare");
        Directory.CreateDirectory(networkSharePath);

        // Split the document by sections.
        for (int i = 0; i < loadedDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document splitDoc = new Document();
            // Remove the default empty section.
            splitDoc.RemoveAllChildren();

            // Import the current section from the source document.
            NodeImporter importer = new NodeImporter(loadedDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(loadedDoc.Sections[i], true);

            // Append the imported section to the new document.
            splitDoc.AppendChild(importedSection);

            // Save the split part.
            string partPath = Path.Combine(networkSharePath, $"Section_{i + 1}.docx");
            splitDoc.Save(partPath);
        }

        // Validate that each expected file exists.
        for (int i = 0; i < loadedDoc.Sections.Count; i++)
        {
            string partPath = Path.Combine(networkSharePath, $"Section_{i + 1}.docx");
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Expected split file not found: {partPath}");
        }

        // Optional cleanup (commented out).
        // Directory.Delete(tempFolder, true);
        // Directory.Delete(networkSharePath, true);
    }
}
