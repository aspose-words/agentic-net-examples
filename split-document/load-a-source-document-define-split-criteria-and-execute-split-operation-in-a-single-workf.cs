using System;
using System.IO;
using Aspose.Words;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Create a sample source document with multiple sections.
        const string sourcePath = "Source.docx";
        CreateSampleDocument(sourcePath);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Prepare output folder.
        string outputFolder = "SplitOutput";
        Directory.CreateDirectory(outputFolder);

        // Split the document by sections.
        int sectionIndex = 1;
        foreach (Section section in sourceDoc.Sections)
        {
            // Create a new empty document for the current section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // remove the default empty section

            // Import the section from the source document into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = Path.Combine(outputFolder, $"Section_{sectionIndex}.docx");
            splitDoc.Save(splitPath);

            // Verify that the file was created.
            if (!File.Exists(splitPath))
                throw new InvalidOperationException($"Failed to create split file: {splitPath}");

            sectionIndex++;
        }

        // Verify the expected number of split files.
        int expectedFiles = sourceDoc.Sections.Count;
        int actualFiles = Directory.GetFiles(outputFolder, "Section_*.docx").Length;
        if (actualFiles != expectedFiles)
            throw new InvalidOperationException($"Expected {expectedFiles} split files, but found {actualFiles}.");
    }

    // Helper method to create a sample document with three sections.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            // Add a header to the current section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln($"Header for Section {i}");

            // Move to the main body and write section content.
            builder.MoveToDocumentEnd();
            builder.Writeln($"This is the content of Section {i}.");

            // Insert a section break (new page) after each section except the last.
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        doc.Save(filePath);
    }
}
