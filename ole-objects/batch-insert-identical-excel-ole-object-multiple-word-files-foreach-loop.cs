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
            // Path to the Excel file that will be embedded as an OLE object.
            string excelFilePath = @"C:\Data\Sample.xlsx";

            // List of Word documents to process.
            List<string> wordFiles = new List<string>
            {
                @"C:\Docs\Report1.docx",
                @"C:\Docs\Report2.docx",
                @"C:\Docs\Report3.docx"
            };

            // Folder where the modified documents will be saved.
            string outputFolder = @"C:\Docs\Processed";

            // Ensure the output folder exists.
            Directory.CreateDirectory(outputFolder);

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
        }
    }
}
