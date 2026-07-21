using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three sections.
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

        // Save the original document.
        string originalPath = "Original.docx";
        sourceDoc.Save(originalPath);

        // Split the document by sections.
        List<Document> splitList = new List<Document>();
        foreach (Section sec in sourceDoc.Sections)
        {
            Document part = new Document();
            part.RemoveAllChildren(); // Remove the default empty section.

            // Import the section from the source document into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, part, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sec, true);
            part.AppendChild(importedSection);
            splitList.Add(part);
        }

        Document[] splitDocs = splitList.ToArray();

        // Save each split part.
        for (int i = 0; i < splitDocs.Length; i++)
        {
            string partPath = $"Part{i}.docx";
            splitDocs[i].Save(partPath);
        }

        // Validate that split parts were saved.
        for (int i = 0; i < splitDocs.Length; i++)
        {
            string partPath = $"Part{i}.docx";
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split part not found: {partPath}");
        }

        // Merge selected parts (e.g., first and third sections) into a new document.
        Document mergedDoc = new Document();
        mergedDoc.RemoveAllChildren(); // Start with an empty document.

        // Define which parts to merge (indices).
        int[] partsToMerge = { 0, 2 }; // first and third sections.

        foreach (int index in partsToMerge)
        {
            if (index < 0 || index >= splitDocs.Length)
                throw new ArgumentOutOfRangeException($"Invalid part index: {index}");

            Document part = splitDocs[index];
            foreach (Section sec in part.Sections)
            {
                NodeImporter importer = new NodeImporter(part, mergedDoc, ImportFormatMode.KeepSourceFormatting);
                Section importedSection = (Section)importer.ImportNode(sec, true);
                mergedDoc.AppendChild(importedSection);
            }
        }

        // Save the merged document.
        string mergedPath = "Merged.docx";
        mergedDoc.Save(mergedPath);

        // Validate merged document exists.
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException($"Merged document not found: {mergedPath}");

        // Simple console output to indicate success (no user interaction required).
        Console.WriteLine("Document splitting and merging completed successfully.");
    }
}
