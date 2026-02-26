using System;
using System.IO;
using Aspose.Words;

class SplitDocumentBySections
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\SourceDocument.docx";

        // Directory where the split section files will be saved.
        string outputDir = @"C:\Docs\SplitSections";

        // Ensure the output directory exists.
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // Load the source document (uses the provided Document(string) constructor).
        Document sourceDoc = new Document(inputFile);

        // Iterate through each section in the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new blank document (uses the parameterless Document() constructor).
            Document sectionDoc = new Document();

            // Remove the default empty section that a blank document contains.
            sectionDoc.Sections.Clear();

            // Clone the current section from the source document.
            Section clonedSection = (Section)sourceDoc.Sections[i].Clone(true);

            // Add the cloned section to the new document.
            sectionDoc.Sections.Add(clonedSection);

            // Build the output file name for this section.
            string outputFile = Path.Combine(outputDir, $"Section_{i + 1}.docx");

            // Save the document (uses the provided Document.Save(string) method).
            sectionDoc.Save(outputFile);
        }

        Console.WriteLine("Document split into sections successfully.");
    }
}
