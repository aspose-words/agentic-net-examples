using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

namespace AsposeWordsBatchProcessing
{
    public class BatchClearer
    {
        /// <summary>
        /// Clears the entire content of each document specified in <paramref name="inputFiles"/>
        /// and saves the cleared document to <paramref name="outputFolder"/> preserving the original file name.
        /// </summary>
        /// <param name="inputFiles">Full paths of the source documents.</param>
        /// <param name="outputFolder">Folder where cleared documents will be saved.</param>
        public void ClearDocuments(IEnumerable<string> inputFiles, string outputFolder)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            foreach (string inputPath in inputFiles)
            {
                // Load the document from file.
                Document doc = new Document(inputPath);

                // Delete all characters in the document's main range, effectively clearing its content.
                doc.Range.Delete();

                // Build the output file path.
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));

                // Save the cleared document. The format is inferred from the file extension.
                doc.Save(outputPath);
            }
        }
    }

    // Example usage:
    class Program
    {
        static void Main()
        {
            // List of document paths to process.
            var files = new List<string>
            {
                @"C:\Docs\Document1.docx",
                @"C:\Docs\Document2.docx",
                @"C:\Docs\Document3.docx"
            };

            // Destination folder for cleared documents.
            string outputFolder = @"C:\ClearedDocs";

            // Perform the batch clear operation.
            var clearer = new BatchClearer();
            clearer.ClearDocuments(files, outputFolder);

            Console.WriteLine("Batch clearing completed.");
        }
    }
}
