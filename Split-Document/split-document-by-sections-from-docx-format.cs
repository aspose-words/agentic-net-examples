using System;
using System.IO;
using Aspose.Words;

class SplitBySections
{
    static void Main()
    {
        // Load the source DOCX document.
        string inputPath = "input.docx";
        Document sourceDoc = new Document(inputPath);

        // Ensure the output directory exists.
        string outputFolder = "OutputSections";
        Directory.CreateDirectory(outputFolder);

        // Iterate through each section in the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document partDoc = new Document();

            // Remove the default empty section that the constructor adds.
            partDoc.Sections.Clear();

            // Import the current section from the source document into the new document.
            Section srcSection = sourceDoc.Sections[i];
            Node importedSection = partDoc.ImportNode(srcSection, true);

            // Append the imported section to the new document.
            partDoc.AppendChild(importedSection);

            // Save the split document using a sequential file name.
            string outputPath = Path.Combine(outputFolder, $"Section_{i + 1}.docx");
            partDoc.Save(outputPath);
        }
    }
}
