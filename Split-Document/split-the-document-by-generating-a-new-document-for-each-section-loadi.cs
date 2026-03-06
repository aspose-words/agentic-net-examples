using System;
using System.IO;
using Aspose.Words;

class SplitDocumentBySections
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Directory where the split section documents will be saved.
        string outputDir = @"C:\Docs\SplitSections";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the source document (DOCX format).
        Document sourceDoc = new Document(sourcePath);

        // Iterate through each section in the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Clone the current section to obtain a deep copy.
            Section clonedSection = sourceDoc.Sections[i].Clone();

            // Create a new blank document.
            Document sectionDoc = new Document();

            // Remove the default empty section that comes with a new document.
            sectionDoc.RemoveAllChildren();

            // Append the cloned section as the sole section of the new document.
            sectionDoc.AppendChild(clonedSection);

            // Build the output file name, e.g., Section_1.docx, Section_2.docx, etc.
            string outFile = Path.Combine(outputDir, $"Section_{i + 1}.docx");

            // Save the new document containing only this section.
            sectionDoc.Save(outFile);
        }

        Console.WriteLine("Document split into sections successfully.");
    }
}
