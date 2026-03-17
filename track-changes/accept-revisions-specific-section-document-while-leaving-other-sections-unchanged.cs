using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace RevisionProcessing
{
    class Program
    {
        static void Main()
        {
            // Load the source document.
            Document doc = new Document("Input.docx");

            // Index of the section whose revisions we want to accept.
            // Sections are zero‑based; change this value as needed.
            int targetSectionIndex = 1;

            // Ensure the index is within the document's section count.
            if (targetSectionIndex >= 0 && targetSectionIndex < doc.Sections.Count)
            {
                // Get the specific section.
                Section targetSection = doc.Sections[targetSectionIndex];

                // Iterate over all revisions that exist in this section's range.
                // The Revisions collection is live, so accepting a revision removes it from the collection.
                foreach (Revision rev in targetSection.Range.Revisions)
                {
                    rev.Accept();
                }
            }

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
