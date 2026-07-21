using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Folder containing the source Word documents.
        string sourceFolder = "InputDocs";

        // Folder where the modified documents will be saved.
        string outputFolder = "OutputDocs";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Verify that the source folder exists; if not, exit gracefully.
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder \"{sourceFolder}\" does not exist. No files to process.");
            return;
        }

        // Path to the Excel file that will be embedded as an OLE object.
        // Ensure this file exists before running the program.
        string excelFilePath = "Sample.xlsx";

        if (!File.Exists(excelFilePath))
        {
            Console.WriteLine($"Excel file \"{excelFilePath}\" not found. Cannot embed OLE object.");
            return;
        }

        // Process each .docx file in the source folder.
        foreach (string docPath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the existing Word document.
            Document doc = new Document(docPath);

            // Create a DocumentBuilder to modify the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph describing the OLE object.
            builder.Writeln("Inserted Excel OLE object:");

            // Embed the Excel file as an OLE object (embedded, not linked, not as an icon).
            builder.InsertOleObject(excelFilePath, false, false, null);

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(docPath));
            doc.Save(outputPath);
        }

        Console.WriteLine("Processing completed.");
    }
}
