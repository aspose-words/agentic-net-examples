using System;
using System.IO;
using Aspose.Words;

class SplitDocumentBySection
{
    static void Main()
    {
        // Path to the source DOCX document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Folder where the split section documents will be saved.
        string outputFolder = @"C:\Docs\SplitSections";

        // Ensure the output folder exists.
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // Load the source document (DOCX format).
        Document sourceDoc = new Document(inputPath);

        // Iterate through each section in the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Clone the current section to create an independent copy.
            Section clonedSection = (Section)sourceDoc.Sections[i].Clone(true);

            // Create a new blank document.
            Document sectionDoc = new Document();

            // Remove any default nodes that Aspose.Words adds to a new document.
            sectionDoc.RemoveAllChildren();

            // Append the cloned section as the sole section of the new document.
            sectionDoc.AppendChild(clonedSection);

            // Build the output file name, e.g., "Section_1.docx", "Section_2.docx", etc.
            string outputPath = Path.Combine(outputFolder, $"Section_{i + 1}.docx");

            // Save the document containing only this section.
            sectionDoc.Save(outputPath);
        }
    }
}
