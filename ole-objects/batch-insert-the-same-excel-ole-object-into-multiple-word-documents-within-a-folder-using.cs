using System;
using System.IO;
using Aspose.Words;

public class BatchOleInserter
{
    public static void Main()
    {
        // Folder containing the Word documents to process.
        // Adjust these paths as needed; they can be absolute or relative.
        string inputFolder = @"C:\Docs\Input";
        string outputFolder = @"C:\Docs\Output";
        string excelFilePath = @"C:\Data\Sample.xlsx";

        // Verify that the input folder exists; if not, inform the user and exit gracefully.
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all .docx files in the input folder.
        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx");

        foreach (string docPath in docFiles)
        {
            // Load the existing Word document.
            Document doc = new Document(docPath);

            // Create a DocumentBuilder for the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move to the end of the document to insert the OLE object.
            builder.MoveToDocumentEnd();

            // Optional: add a paragraph before the OLE object.
            builder.Writeln("Embedded Excel workbook:");

            // Insert the Excel file as an embedded OLE object (not as an icon).
            // Using the overload: InsertOleObject(string fileName, bool isLinked, bool asIcon, Stream presentation)
            builder.InsertOleObject(excelFilePath, false, false, null);

            // Determine the output file path (same file name, different folder).
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(docPath));

            // Save the modified document.
            doc.Save(outputPath);
        }

        Console.WriteLine("Processing completed.");
    }
}
