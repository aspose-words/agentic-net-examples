using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;

namespace OleBatchInsertExample
{
    public class Program
    {
        // Path to the Excel file that will be embedded as an OLE object.
        private static readonly string ExcelFilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Sample.xlsx"));

        // Folder containing the Word documents to be processed.
        private static readonly string InputFolder = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Input"));

        // Folder where the modified documents will be saved.
        private static readonly string OutputFolder = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Output"));

        public static void Main()
        {
            // Verify that the Excel file exists.
            if (!File.Exists(ExcelFilePath))
            {
                Console.WriteLine($"Excel file not found: {ExcelFilePath}");
                return;
            }

            // Ensure the input and output directories exist.
            Directory.CreateDirectory(InputFolder);
            Directory.CreateDirectory(OutputFolder);

            // Collect all .docx files from the input folder.
            List<string> wordFiles = new List<string>(Directory.GetFiles(InputFolder, "*.docx"));

            // If there are no files, inform the user and exit gracefully.
            if (wordFiles.Count == 0)
            {
                Console.WriteLine($"No .docx files found in: {InputFolder}");
                return;
            }

            foreach (string wordFilePath in wordFiles)
            {
                // Load the existing Word document.
                Document doc = new Document(wordFilePath);

                // Create a DocumentBuilder for the loaded document.
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Move the cursor to the end of the document.
                builder.MoveToDocumentEnd();

                // Insert a paragraph break before the OLE object for readability.
                builder.Writeln();

                // Insert the Excel OLE object (embedded, not as an icon, no custom presentation image).
                // Overload used: InsertOleObject(string fileName, bool isLinked, bool asIcon, Stream presentation)
                builder.InsertOleObject(ExcelFilePath, false, false, null);

                // Determine the output file name (preserve original name, place in output folder).
                string outputFilePath = Path.Combine(OutputFolder, Path.GetFileName(wordFilePath));

                // Save the modified document.
                doc.Save(outputFilePath);
            }

            Console.WriteLine("OLE objects inserted successfully.");
        }
    }
}
