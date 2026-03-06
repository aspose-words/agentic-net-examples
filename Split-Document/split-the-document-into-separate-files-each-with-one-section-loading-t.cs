using System;
using Aspose.Words;

class SplitDocumentBySection
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("Source.docx");

        // Iterate through each section in the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document partDoc = new Document();

            // Clone the current section to keep the original unchanged.
            Section clonedSection = (Section)sourceDoc.Sections[i].Clone(true);

            // Remove the default empty section that a new document contains.
            partDoc.Sections.Clear();

            // Add the cloned section to the new document.
            partDoc.Sections.Add(clonedSection);

            // Save the new document; each file will contain a single section.
            string outFileName = $"Section_{i + 1}.docx";
            partDoc.Save(outFileName);
        }
    }
}
