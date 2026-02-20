using System;
using Aspose.Words;

class SplitDocumentBySections
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("InputDocument.docx");

        // Iterate through each section in the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Get the current section.
            Section srcSection = sourceDoc.Sections[i];

            // Create a new empty document.
            Document partDoc = new Document();

            // Remove the default empty section that a new document contains.
            partDoc.RemoveAllChildren();

            // Prepare a NodeImporter to import nodes from the source to the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);

            // Import the section (deep clone) into the new document.
            Section importedSection = (Section)importer.ImportNode(srcSection, true);

            // Append the imported section to the new document.
            partDoc.AppendChild(importedSection);

            // Save the split part as a separate DOCX file.
            string outputFileName = $"Section_{i + 1}.docx";
            partDoc.Save(outputFileName);
        }
    }
}
