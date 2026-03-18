using System;
using System.IO;
using Aspose.Words;

namespace RevisionBatchProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the documents to process.
            // Adjust the path as needed.
            string inputFolder = @"C:\Docs\ToProcess";

            // Ensure the folder exists.
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Folder not found: {inputFolder}");
                return;
            }

            // Process all Word documents in the folder (including subfolders if desired).
            string[] files = Directory.GetFiles(inputFolder, "*.docx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                try
                {
                    // Load the document from the file (uses the Document(string) constructor).
                    Document doc = new Document(filePath);

                    // Accept all tracked changes in the document.
                    doc.AcceptAllRevisions();

                    // Save the document back to the same location, overwriting the original file.
                    doc.Save(filePath);
                    
                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing remaining files.
                    Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch revision acceptance completed.");
        }
    }
}
