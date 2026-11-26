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
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Content of section {i}.");
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the source document (demonstrates load workflow).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document from file.
        Document doc = new Document(sourcePath);

        // Ensure the output folder exists.
        string outputFolder = "SplitParts";
        Directory.CreateDirectory(outputFolder);

        // Split the document by sections manually.
        IList<Document> splitDocs = new List<Document>();
        foreach (Section section in doc.Sections)
        {
            // Create a new empty document.
            Document partDoc = new Document();
            // Remove any existing content (the default empty section).
            partDoc.RemoveAllChildren();

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(doc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            partDoc.AppendChild(importedSection);

            splitDocs.Add(partDoc);
        }

        // Define the callback that saves each split part.
        Action<Document, string, int, SaveOptions> documentPartSavingCallback = (partDoc, partName, partIndex, saveOptions) =>
        {
            // Save each split part using the provided file name.
            partDoc.Save(partName);
        };

        // Iterate over the split collection and save each part using the callback.
        for (int i = 0; i < splitDocs.Count; i++)
        {
            string partFileName = $"Part_{i + 1}.docx";
            string partPath = Path.Combine(outputFolder, partFileName);
            documentPartSavingCallback(splitDocs[i], partPath, i, null);
        }

        // Verify that each expected file was created.
        for (int i = 0; i < splitDocs.Count; i++)
        {
            string expectedPath = Path.Combine(outputFolder, $"Part_{i + 1}.docx");
            if (!File.Exists(expectedPath))
                throw new InvalidOperationException($"Expected split file not found: {expectedPath}");
        }
    }
}
