using System;
using System.IO;
using Aspose.Words;

public class BatchOleInsert
{
    public static void Main()
    {
        // Folder containing the Word documents to process.
        // Use a relative path that is more likely to exist, or adjust as needed.
        string inputFolder = Path.Combine(Environment.CurrentDirectory, "Input");
        // Folder where the modified documents will be saved.
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "Output");
        // Path to the Excel file that will be inserted as an OLE object.
        string excelFilePath = Path.Combine(Environment.CurrentDirectory, "Sample.xlsx");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Verify that the input folder exists; if not, exit gracefully.
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Verify that the Excel file exists; if not, exit gracefully.
        if (!File.Exists(excelFilePath))
        {
            Console.WriteLine($"Excel file not found: {excelFilePath}");
            return;
        }

        // Process each .docx file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the existing Word document.
            Document doc = new Document(docPath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph describing the OLE object.
            builder.Writeln("Embedded Excel OLE object:");

            // Insert the Excel file as an embedded OLE object (not as an icon).
            // Parameters: fileName, isLinked (false = embed), asIcon (false = show content), presentation (null).
            builder.InsertOleObject(excelFilePath, false, false, null);

            // Determine the output file path (same file name, different folder).
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(docPath));

            // Save the modified document.
            doc.Save(outputPath);
        }

        Console.WriteLine("Processing completed.");
    }
}
