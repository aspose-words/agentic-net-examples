using System;
using Aspose.Words;

class SplitDocumentBySections
{
    static void Main()
    {
        // Load the source DOCX document.
        string sourcePath = "InputDocument.docx";
        Document sourceDoc = new Document(sourcePath);

        // Iterate through each section in the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Clone the current section.
            Section clonedSection = sourceDoc.Sections[i].Clone();

            // Create a new blank document.
            Document newDoc = new Document();

            // Remove the default empty section that Aspose.Words creates.
            newDoc.Sections.Clear();

            // Append the cloned section to the new document.
            newDoc.Sections.Add(clonedSection);

            // Save the new document containing only this section.
            string outputPath = $"Section_{i + 1}.docx";
            newDoc.Save(outputPath);
        }
    }
}
