using System;
using Aspose.Words;

class SplitDocumentBySections
{
    static void Main()
    {
        // Load the source DOCX file.
        Document sourceDoc = new Document("InputDocument.docx");

        // Iterate through each section in the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new blank document.
            Document partDoc = new Document();

            // Remove the default empty section that a new document contains.
            partDoc.Sections.Clear();

            // Clone the current section from the source document (deep clone) and cast to Section.
            Section clonedSection = (Section)sourceDoc.Sections[i].Clone(true);

            // Add the cloned section to the new document.
            partDoc.Sections.Add(clonedSection);

            // Save the new document; each file will contain a single original section.
            string outputPath = $"Section_{i + 1}.docx";
            partDoc.Save(outputPath);
        }
    }
}
