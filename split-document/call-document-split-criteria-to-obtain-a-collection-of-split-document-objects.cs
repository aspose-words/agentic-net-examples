using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with three sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - Paragraph 1");

        // Insert a section break to start Section 3.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3 - Paragraph 1");

        // Split the document into separate documents, one per section.
        List<Document> splitDocuments = new List<Document>();
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document that will hold the split part.
            Document partDoc = new Document();

            // Import the current section from the source document into the new document.
            // ImportNode handles cloning and ensures the node belongs to the destination document.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true, ImportFormatMode.KeepSourceFormatting);
            partDoc.AppendChild(importedSection);

            splitDocuments.Add(partDoc);
        }

        // Save each split document to the local file system and verify the files exist.
        for (int i = 0; i < splitDocuments.Count; i++)
        {
            string fileName = $"SplitPart_{i + 1}.docx";
            splitDocuments[i].Save(fileName);

            if (!File.Exists(fileName))
                throw new InvalidOperationException($"Failed to create split file: {fileName}");
        }

        // Optional: output a simple confirmation.
        Console.WriteLine($"Successfully split document into {splitDocuments.Count} parts.");
    }
}
