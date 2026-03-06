using System;
using System.IO;
using Aspose.Words;

namespace DocumentSplittingExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourceFilePath = @"C:\Docs\SourceDocument.docx";

            // Load the source document. This uses the Document(string) constructor.
            Document sourceDocument = new Document(sourceFilePath);

            // Ensure the document contains at least one section.
            if (sourceDocument.Sections.Count == 0)
                throw new InvalidOperationException("The source document does not contain any sections.");

            // Directory where the split parts will be saved.
            string outputDirectory = @"C:\Docs\SplitParts";
            Directory.CreateDirectory(outputDirectory);

            // Iterate through each section and save it as an individual document.
            for (int i = 0; i < sourceDocument.Sections.Count; i++)
            {
                // Create a new blank document.
                Document partDocument = new Document();

                // Clone the current section from the source document.
                // The Clone method creates a deep copy of the section node.
                Node clonedSection = sourceDocument.Sections[i].Clone(true);

                // Append the cloned section to the new document.
                // AppendChild adds the node to the end of the child nodes collection.
                partDocument.AppendChild(clonedSection);

                // Build the output file name, e.g., "SourceDocument_part_1.docx".
                string partFileName = Path.GetFileNameWithoutExtension(sourceFilePath) +
                                      $"_part_{i + 1}{Path.GetExtension(sourceFilePath)}";

                string partFilePath = Path.Combine(outputDirectory, partFileName);

                // Save the part document. This uses the Document.Save(string) method.
                partDocument.Save(partFilePath);
            }

            Console.WriteLine("Document split into sections successfully.");
        }
    }
}
