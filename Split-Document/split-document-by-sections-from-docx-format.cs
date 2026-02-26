using System;
using System.IO;
using Aspose.Words;

namespace DocumentSplitBySections
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document sourceDoc = new Document("input.docx");

            // Iterate through each section in the source document.
            for (int i = 0; i < sourceDoc.Sections.Count; i++)
            {
                // Create a new blank document (lifecycle rule: create).
                Document partDoc = new Document();

                // Remove any default nodes that a blank document contains.
                partDoc.RemoveAllChildren();

                // Import the current section from the source document into the new document.
                // ImportNode preserves formatting (importNode rule).
                Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true);

                // Append the imported section to the new document.
                partDoc.AppendChild(importedSection);

                // Save the split part (lifecycle rule: save).
                string outputPath = $"Section_{i + 1}.docx";
                partDoc.Save(outputPath);
            }
        }
    }
}
