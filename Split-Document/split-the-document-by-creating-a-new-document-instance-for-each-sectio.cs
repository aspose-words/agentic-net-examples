using System;
using System.IO;
using Aspose.Words;

namespace DocumentSplitter
{
    /// <summary>
    /// Provides functionality to split a DOCX document into separate documents,
    /// one for each section.
    /// </summary>
    public static class SectionSplitter
    {
        /// <summary>
        /// Splits the source document into individual section documents.
        /// </summary>
        /// <param name="sourceFilePath">Full path to the source DOCX file.</param>
        /// <param name="outputFolderPath">Folder where the split documents will be saved.</param>
        public static void SplitDocumentBySections(string sourceFilePath, string outputFolderPath)
        {
            // Load the source document using the Document(string) constructor.
            Document sourceDoc = new Document(sourceFilePath);

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolderPath);

            // Iterate through each section in the source document.
            for (int i = 0; i < sourceDoc.Sections.Count; i++)
            {
                // Create a new blank document.
                Document partDoc = new Document();

                // Remove the default empty section that a new Document contains.
                partDoc.RemoveAllChildren();

                // Clone the current section from the source document.
                // The Clone method returns a Node, so cast it to Section.
                Section clonedSection = (Section)sourceDoc.Sections[i].Clone(true);

                // Append the cloned section to the new document.
                partDoc.AppendChild(clonedSection);

                // Build a filename for the part document, e.g., "Section_1.docx".
                string partFileName = $"Section_{i + 1}.docx";
                string partFilePath = Path.Combine(outputFolderPath, partFileName);

                // Save the part document using the Document.Save(string) method.
                partDoc.Save(partFilePath);
            }
        }

        // Example usage.
        public static void Main()
        {
            // Path to the source DOCX file.
            string sourcePath = @"C:\Docs\SourceDocument.docx";

            // Folder where the split sections will be saved.
            string outputPath = @"C:\Docs\SplitSections";

            // Perform the split operation.
            SplitDocumentBySections(sourcePath, outputPath);

            Console.WriteLine("Document split into sections successfully.");
        }
    }
}
