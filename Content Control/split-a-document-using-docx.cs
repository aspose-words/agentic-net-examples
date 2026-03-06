using System;
using System.IO;
using Aspose.Words;

class SplitDocx
{
    static void Main()
    {
        // Load the source DOCX document.
        Document source = new Document("Input.docx");

        // Directory where the split parts will be saved.
        string outputFolder = "SplitParts";
        Directory.CreateDirectory(outputFolder);

        // Iterate through each section of the source document.
        for (int i = 0; i < source.Sections.Count; i++)
        {
            // Create a new blank document that will hold a single section.
            Document part = new Document();

            // Remove the default empty section that Aspose.Words creates.
            part.RemoveAllChildren();

            // Import the current section from the source document into the new document.
            // ImportNode clones the node and adjusts any references (styles, images, etc.).
            Section importedSection = (Section)part.ImportNode(source.Sections[i], true);

            // Append the imported section as the sole section of the new document.
            part.AppendChild(importedSection);

            // Build the output file name (e.g., Part_1.docx, Part_2.docx, ...).
            string outFile = Path.Combine(outputFolder, $"Part_{i + 1}.docx");

            // Save the split part as a DOCX file.
            part.Save(outFile);
        }
    }
}
