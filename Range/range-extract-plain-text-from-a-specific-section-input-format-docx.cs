using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractSectionPlainText
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the extracted plain‑text will be saved.
        string outputPath = @"C:\Docs\SectionText.txt";

        // Index of the section to extract (0‑based).
        int sectionIndex = 2; // example: third section

        // Load the source document (lifecycle: load).
        Document sourceDoc = new Document(inputPath);

        // Validate the requested section index.
        if (sectionIndex < 0 || sectionIndex >= sourceDoc.Sections.Count)
            throw new ArgumentOutOfRangeException(nameof(sectionIndex), "Section index is out of range.");

        // Get the required section.
        Section targetSection = sourceDoc.Sections[sectionIndex];

        // Create a new blank document (lifecycle: create).
        Document extractedDoc = new Document();

        // Remove the default empty section that comes with a new document.
        extractedDoc.RemoveAllChildren();

        // Import the target section into the new document.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
        Section importedSection = (Section)importer.ImportNode(targetSection, true);
        extractedDoc.AppendChild(importedSection);

        // Save the new document as plain text (lifecycle: save).
        extractedDoc.Save(outputPath, SaveFormat.Text);
    }
}
