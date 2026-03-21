using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

namespace OleObjectBatchInsert
{
    class Program
    {
        static void Main()
        {
            // Base directory for all test files (relative to the executable location).
            string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles");
            string excelFilePath = Path.Combine(baseDir, "Sample.xlsx");
            string wordFilesDir = Path.Combine(baseDir, "WordDocs");
            string outputFolder = Path.Combine(baseDir, "Processed");

            // Ensure directories exist.
            Directory.CreateDirectory(baseDir);
            Directory.CreateDirectory(wordFilesDir);
            Directory.CreateDirectory(outputFolder);

            // Create a dummy Excel file if it does not exist.
            if (!File.Exists(excelFilePath))
            {
                // Write minimal content; it does not need to be a valid workbook for the demo.
                File.WriteAllText(excelFilePath, "Dummy Excel content");
            }

            // Prepare a list of Word document paths.
            List<string> wordFiles = new List<string>();
            for (int i = 1; i <= 3; i++)
            {
                string wordPath = Path.Combine(wordFilesDir, $"Report{i}.docx");
                if (!File.Exists(wordPath))
                {
                    // Create a simple Word document with some placeholder text.
                    Document tempDoc = new Document();
                    DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                    tempBuilder.Writeln($"This is Report {i}.");
                    tempDoc.Save(wordPath);
                }
                wordFiles.Add(wordPath);
            }

            // Process each Word document: insert the Excel OLE object and save the result.
            foreach (string wordFilePath in wordFiles)
            {
                // Load the existing Word document.
                Document doc = new Document(wordFilePath);

                // Create a DocumentBuilder positioned at the end of the document.
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveToDocumentEnd();

                // Insert a paragraph break before the OLE object for readability.
                builder.InsertParagraph();

                // Insert the Excel OLE object.
                // Parameters: file name, isLinked (false = embed), asIcon (false = show content), presentation (null = default image).
                builder.InsertOleObject(excelFilePath, false, false, null);

                // Determine the output file name.
                string fileName = Path.GetFileNameWithoutExtension(wordFilePath);
                string outputPath = Path.Combine(outputFolder, $"{fileName}_WithOle.docx");

                // Save the modified document.
                doc.Save(outputPath);
            }

            Console.WriteLine("Processing completed. Check the 'Processed' folder under " + baseDir);
        }
    }
}
