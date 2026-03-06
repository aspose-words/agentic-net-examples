using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentSectionSplitter
{
    static void Main()
    {
        // Load the source DOCX document.
        string sourcePath = @"Input.docx";
        Document sourceDoc = new Document(sourcePath);

        // Iterate through each section in the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document that will contain a single section.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // Ensure the document has no default nodes.

            // Import the current section into the new document.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true, ImportFormatMode.KeepSourceFormatting);
            partDoc.AppendChild(importedSection);

            // Save the split part using a filename that reflects the section index.
            string outFileName = $"Section_{i + 1}.docx";
            partDoc.Save(outFileName, SaveFormat.Docx);
        }
    }
}
