using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path to the Excel file that will be inserted as an OLE object.
        // Use a full path to avoid relative‑path issues when the program is run from another folder.
        string excelFilePath = Path.GetFullPath(Path.Combine("Resources", "Template.xlsx"));

        // Folder containing the source Word documents.
        string inputFolder = Path.GetFullPath("InputDocs");

        // Folder where the modified documents will be saved.
        string outputFolder = Path.GetFullPath("OutputDocs");

        // Verify that the required folders exist.
        if (!File.Exists(excelFilePath))
        {
            Console.WriteLine($"Excel template not found: {excelFilePath}");
            return;
        }

        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all .docx files from the input folder.
        string[] wordFiles = Directory.GetFiles(inputFolder, "*.docx");

        foreach (string wordFilePath in wordFiles)
        {
            // Load the existing Word document.
            Document doc = new Document(wordFilePath);

            // Create a DocumentBuilder to modify the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph describing the OLE object.
            builder.Writeln("Inserted Excel OLE object:");

            // Insert the Excel file as an embedded OLE object (not as an icon).
            // Parameters: file name, isLinked = false (embed), asIcon = false (show content), presentation = null.
            builder.InsertOleObject(excelFilePath, false, false, null);

            // Build the output file name.
            string outputFileName = Path.GetFileNameWithoutExtension(wordFilePath) + "_Ole.docx";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the modified document.
            doc.Save(outputPath);
        }

        Console.WriteLine("Processing completed.");
    }
}
