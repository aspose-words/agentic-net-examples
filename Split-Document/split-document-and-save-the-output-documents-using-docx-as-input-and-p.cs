using System;
using Aspose.Words;

namespace DocumentSplitter
{
    class Program
    {
        static void Main()
        {
            // Input DOCX file path.
            string inputPath = @"C:\Docs\InputDocument.docx";

            // Load the source document (DOCX).
            Document sourceDoc = new Document(inputPath);

            // Iterate through each section of the source document.
            for (int i = 0; i < sourceDoc.Sections.Count; i++)
            {
                // Create a new empty document.
                Document partDoc = new Document();

                // Clone the current section (deep clone) and add it to the new document.
                // The cloned section already contains its body, paragraphs, etc.
                Section clonedSection = (Section)sourceDoc.Sections[i].Clone(true);
                partDoc.AppendChild(clonedSection);

                // Build an output file name for the PDF part.
                string outputPath = $@"C:\Docs\Output_Part_{i + 1}.pdf";

                // Save the part as PDF. The format is inferred from the file extension.
                partDoc.Save(outputPath);
            }
        }
    }
}
