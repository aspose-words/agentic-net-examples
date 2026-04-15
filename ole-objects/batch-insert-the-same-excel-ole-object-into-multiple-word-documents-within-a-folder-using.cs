using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder containing the source Word documents.
        string inputFolder = Path.Combine(Environment.CurrentDirectory, "InputDocs");
        // Folder where the modified documents will be saved.
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "OutputDocs");
        // Path to the Excel file that will be inserted as an OLE object.
        string excelFilePath = Path.Combine(Environment.CurrentDirectory, "Template.xlsx");

        // Ensure the input and output directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Verify that the Excel template exists before proceeding.
        if (!File.Exists(excelFilePath))
        {
            Console.WriteLine($"Excel template not found: {excelFilePath}");
            return;
        }

        // Get all .docx files in the input folder.
        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx", SearchOption.TopDirectoryOnly);

        // If there are no documents, inform the user and exit gracefully.
        if (docFiles.Length == 0)
        {
            Console.WriteLine($"No .docx files found in: {inputFolder}");
            return;
        }

        foreach (string docPath in docFiles)
        {
            // Load the existing Word document.
            Document doc = new Document(docPath);

            // Create a DocumentBuilder to modify the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the Excel OLE object at the current cursor position.
            // Parameters: file name, isLinked (false = embed), asIcon (false = show content), presentation (null = default icon).
            builder.InsertOleObject(excelFilePath, false, false, null);

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(docPath));
            doc.Save(outputPath);
        }

        Console.WriteLine("Batch insertion completed successfully.");
    }
}
