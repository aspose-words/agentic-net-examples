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
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Warning: Input file not found and will be skipped: {inputPath}");
                    continue;
                }

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

    class Program
    {
        static void Main()
        {
            // Create a temporary folder for sample input documents.
            string tempInputFolder = Path.Combine(Path.GetTempPath(), "AsposeSampleInput");
            Directory.CreateDirectory(tempInputFolder);

            // Create a few simple documents to work with.
            var files = new List<string>();
            for (int i = 1; i <= 3; i++)
            {
                string filePath = Path.Combine(tempInputFolder, $"Document{i}.docx");
                // Create a new empty document with a single paragraph of sample text.
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"Sample content for Document{i}");
                doc.Save(filePath);
                files.Add(filePath);
            }

            // Destination folder for cleared documents.
            string outputFolder = Path.Combine(Path.GetTempPath(), "AsposeClearedDocs");
            Directory.CreateDirectory(outputFolder);

            // Perform the batch clear operation.
            var clearer = new BatchClearer();
            clearer.ClearDocuments(files, outputFolder);

            Console.WriteLine("Batch clearing completed.");
            Console.WriteLine($"Input files located at: {tempInputFolder}");
            Console.WriteLine($"Cleared files saved at: {outputFolder}");
        }
    }
}
