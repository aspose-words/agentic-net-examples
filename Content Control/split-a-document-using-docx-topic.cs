using System;
using System.IO;
using Aspose.Words;

namespace DocumentSplitter
{
    /// <summary>
    /// Demonstrates how to split a DOCX document into separate DOCX files,
    /// each containing a single section from the original document.
    /// </summary>
    public static class DocxSectionSplitter
    {
        /// <summary>
        /// Splits the input DOCX file into multiple DOCX files, one per section.
        /// </summary>
        /// <param name="inputFilePath">Full path to the source DOCX document.</param>
        /// <param name="outputFolderPath">Folder where the split documents will be saved.</param>
        public static void SplitDocumentBySection(string inputFilePath, string outputFolderPath)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolderPath))
                Directory.CreateDirectory(outputFolderPath);

            // Load the source document using the provided Document constructor.
            Document sourceDoc = new Document(inputFilePath);

            // Iterate through each section in the source document.
            for (int i = 0; i < sourceDoc.Sections.Count; i++)
            {
                // Create a new empty document that will hold the current section.
                Document partDoc = new Document();

                // Import the section from the source document into the new document.
                // ImportNode handles any necessary reference adjustments.
                Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true);

                // Append the imported section as the sole child of the new document.
                partDoc.AppendChild(importedSection);

                // Build the output file name (e.g., Part_1.docx, Part_2.docx, ...).
                string outputFileName = Path.Combine(outputFolderPath, $"Part_{i + 1}.docx");

                // Save the split part using the provided Save method.
                partDoc.Save(outputFileName);
            }
        }

        // Example usage.
        public static void Main()
        {
            // Path to the original DOCX file.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Folder where the split documents will be placed.
            string outputPath = @"C:\Docs\SplitParts";

            // Perform the split operation.
            SplitDocumentBySection(inputPath, outputPath);

            Console.WriteLine("Document split completed.");
        }
    }
}
