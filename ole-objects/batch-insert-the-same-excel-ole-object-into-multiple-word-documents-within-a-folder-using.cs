using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Folder containing the Word documents to process.
        string inputFolder = Path.Combine(Environment.CurrentDirectory, "InputDocs");
        // Path to the Excel file that will be embedded as an OLE object.
        string excelFilePath = Path.Combine(Environment.CurrentDirectory, "SampleData.xlsx");

        // Ensure the input folder exists.
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Verify the Excel file exists.
        if (!File.Exists(excelFilePath))
        {
            Console.WriteLine($"Excel file not found: {excelFilePath}");
            return;
        }

        // Process each .docx file in the folder.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the existing Word document.
            Document doc = new Document(docPath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph describing the OLE object.
            builder.Writeln("Embedded Excel OLE object:");

            // Open the Excel file as a stream and embed it.
            using (FileStream excelStream = File.OpenRead(excelFilePath))
            {
                // Insert the OLE object (embedded, not as an icon, no custom presentation image).
                builder.InsertOleObject(excelStream, "Excel.Sheet", false, null);
            }

            // Save the modified document, overwriting the original file.
            doc.Save(docPath);
        }
    }
}
