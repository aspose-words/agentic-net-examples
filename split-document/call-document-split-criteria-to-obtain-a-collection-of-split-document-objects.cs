using System;
using System.IO;
using Aspose.Words;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Create a sample document with two sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Section 1 - First paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - First paragraph.");

        // Save the source document locally.
        const string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document that will be split.
        Document doc = new Document(sourcePath);

        // Split the document by sections.
        int index = 1;
        foreach (Section section in doc.Sections)
        {
            // Create a new empty document and remove its default empty section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren();

            // Import the current section (including headers/footers) into the new document.
            NodeImporter importer = new NodeImporter(doc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = $"Split_{index}.docx";
            splitDoc.Save(splitPath);

            // Verify that the file was saved.
            if (!File.Exists(splitPath))
                throw new InvalidOperationException($"Split document was not saved: {splitPath}");

            index++;
        }

        // Verify the expected number of split documents.
        int expectedCount = doc.Sections.Count;
        if (index - 1 != expectedCount)
            throw new InvalidOperationException("The number of split documents does not match the expected count.");
    }
}
