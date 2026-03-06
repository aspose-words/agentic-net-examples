using System;
using Aspose.Words;

namespace DocumentSplitter
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document sourceDoc = new Document("SourceDocument.docx");

            // Iterate through each section in the source document.
            for (int i = 0; i < sourceDoc.Sections.Count; i++)
            {
                // Create a new blank document.
                Document splitDoc = new Document();

                // Remove the default empty section that a new document contains.
                splitDoc.Sections.Clear();

                // Clone the current section from the source document.
                Section clonedSection = sourceDoc.Sections[i].Clone();

                // Append the cloned section to the new document.
                splitDoc.Sections.Add(clonedSection);

                // Save the new document. Each file will contain a single section.
                string outputFileName = $"Section_{i + 1}.docx";
                splitDoc.Save(outputFileName);
            }
        }
    }
}
