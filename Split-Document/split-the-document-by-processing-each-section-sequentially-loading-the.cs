using System;
using System.IO;
using Aspose.Words;

namespace DocumentSplitterApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage – replace with your actual paths.
            string sourceFilePath = @"C:\Docs\Source.docx";
            string outputFolderPath = @"C:\Docs\SplitSections";

            SplitDocumentBySections(sourceFilePath, outputFolderPath);
            Console.WriteLine("Document split completed.");
        }

        /// <summary>
        /// Splits a DOCX document into separate files, one per section.
        /// Each section is saved as an individual DOCX file in the specified output folder.
        /// </summary>
        /// <param name="sourceFilePath">Full path to the source DOCX document.</param>
        /// <param name="outputFolderPath">Folder where the split documents will be saved.</param>
        public static void SplitDocumentBySections(string sourceFilePath, string outputFolderPath)
        {
            // Load the source document (DOCX) using the Document(string) constructor.
            Document sourceDoc = new Document(sourceFilePath);

            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolderPath))
                Directory.CreateDirectory(outputFolderPath);

            // Iterate through each section in the source document.
            for (int i = 0; i < sourceDoc.Sections.Count; i++)
            {
                // Create a new blank document.
                Document partDoc = new Document();

                // Remove the automatically created default section.
                partDoc.RemoveAllChildren();

                // Clone the current section (deep copy) and add it to the new document.
                Section clonedSection = (Section)sourceDoc.Sections[i].Clone(true);
                partDoc.AppendChild(clonedSection);

                // Build a filename for the part document.
                string partFileName = $"Section_{i + 1}.docx";
                string partFilePath = Path.Combine(outputFolderPath, partFileName);

                // Save the part document using the Save(string) method.
                partDoc.Save(partFilePath);
            }
        }
    }
}
