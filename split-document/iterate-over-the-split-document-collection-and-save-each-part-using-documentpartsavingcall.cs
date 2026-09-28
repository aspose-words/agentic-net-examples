using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document that contains three sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Content of Section {i}.");
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the source document for reference (optional).
        sourceDoc.Save("Source.docx");

        // Split the source document into separate documents, one per section.
        List<Document> splitDocs = new List<Document>();
        foreach (Section section in sourceDoc.Sections)
        {
            // Create a new empty document that will hold the imported section.
            Document partDoc = new Document();

            // Import the section from the source document into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            partDoc.AppendChild(importedSection);

            splitDocs.Add(partDoc);
        }

        // Save each split document using a deterministic file name.
        for (int i = 0; i < splitDocs.Count; i++)
        {
            string fileName = $"Part_{i + 1}.docx";
            splitDocs[i].Save(fileName);
        }

        // Verify that each expected output file exists.
        for (int i = 1; i <= splitDocs.Count; i++)
        {
            string fileName = $"Part_{i}.docx";
            if (!File.Exists(fileName))
                throw new Exception($"Expected file '{fileName}' was not created.");
        }

        Console.WriteLine("All split parts were saved successfully.");
    }
}
