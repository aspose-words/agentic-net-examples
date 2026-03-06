using System;
using System.IO;
using Aspose.Words;

class SplitDocumentToDocx
{
    static void Main()
    {
        // Load the source document.
        string inputPath = @"MyDir\Source.docx";
        Document sourceDoc = new Document(inputPath);

        // Ensure the output folder exists.
        string outputFolder = @"ArtifactsDir\SplitParts";
        Directory.CreateDirectory(outputFolder);

        // Iterate through each section of the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document for the current part.
            Document partDoc = new Document();

            // Import the current section into the new document.
            // The ImportNode method clones the section and resolves any node references.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true);

            // Remove any default empty section and add the imported one.
            partDoc.Sections.Clear();
            partDoc.Sections.Add(importedSection);

            // Build a unique file name for the part.
            string partFileName = Path.Combine(outputFolder, $"Part_{i + 1}.docx");

            // Save the part as a separate DOCX file.
            partDoc.Save(partFileName);
        }
    }
}
