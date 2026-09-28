using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("Section 1 - First paragraph.");
        builder.Writeln("Section 1 - Second paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - First paragraph.");
        builder.Writeln("Section 2 - Second paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - First paragraph.");
        builder.Writeln("Section 3 - Second paragraph.");

        // Define a folder that represents a network share location.
        // For demonstration purposes we use a local temporary folder.
        string networkSharePath = Path.Combine(Path.GetTempPath(), "NetworkShare");
        Directory.CreateDirectory(networkSharePath);

        // Split the document by sections manually.
        List<Document> splitDocs = new List<Document>();
        foreach (Section section in sourceDoc.Sections)
        {
            // Create a new empty document.
            Document part = new Document();

            // Remove the default empty section that a new Document contains.
            part.Sections.Clear();

            // Import the current section from the source document into the new document.
            Section importedSection = (Section)part.ImportNode(section, true);
            part.Sections.Add(importedSection);

            splitDocs.Add(part);
        }

        // Save each split part to the network share folder.
        for (int i = 0; i < splitDocs.Count; i++)
        {
            string fileName = $"Part_{i + 1}.docx";
            string fullPath = Path.Combine(networkSharePath, fileName);
            splitDocs[i].Save(fullPath, SaveFormat.Docx);

            // Validate that the file was created.
            if (!File.Exists(fullPath))
                throw new InvalidOperationException($"Failed to save split document: {fullPath}");
        }

        // Indicate completion.
        Console.WriteLine($"Successfully split document into {splitDocs.Count} parts at: {networkSharePath}");
    }
}
