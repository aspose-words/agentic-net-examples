using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Create a sample document with three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("Section 1 - First paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - First paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - First paragraph.");

        // Manually split the document by sections.
        List<Document> splitDocs = new List<Document>();
        foreach (Section section in doc.Sections)
        {
            // Create a new empty document for the split part.
            Document part = new Document();

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(doc, part, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);

            // Append the imported section to the new document.
            part.AppendChild(importedSection);
            splitDocs.Add(part);
        }

        // Ensure the split returned the expected number of documents.
        if (splitDocs == null || splitDocs.Count != 3)
            throw new InvalidOperationException("Document splitting did not produce the expected number of parts.");

        // Save each split document to a separate file.
        for (int i = 0; i < splitDocs.Count; i++)
        {
            string fileName = $"SplitPart_{i + 1}.docx";
            splitDocs[i].Save(fileName, SaveFormat.Docx);

            // Verify that the file was created.
            if (!File.Exists(fileName))
                throw new FileNotFoundException($"Failed to create split file: {fileName}");
        }

        // All split documents have been saved successfully.
        Console.WriteLine("Document split completed successfully.");
    }
}
