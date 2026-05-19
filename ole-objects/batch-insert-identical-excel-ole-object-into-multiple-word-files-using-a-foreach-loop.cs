using System;
using System.IO;
using Aspose.Words;

public class BatchOleInsert
{
    public static void Main()
    {
        // Base directory of the running application.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        // Folder that contains the source Word documents and the Excel file to embed.
        string dataDir = Path.Combine(baseDir, "Data");

        // Folder where the modified documents will be saved.
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Full path to the Excel file that will be inserted as an OLE object.
        string excelPath = Path.Combine(dataDir, "Sample.xlsx");

        // Verify that the Excel file exists; if not, abort with a clear message.
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Excel file not found: {excelPath}");
            return;
        }

        // Get all Word documents (*.docx) in the data directory.
        string[] wordFiles = Directory.GetFiles(dataDir, "*.docx");

        foreach (string wordFilePath in wordFiles)
        {
            // Load the existing Word document.
            Document doc = new Document(wordFilePath);

            // Create a DocumentBuilder to modify the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Optional description paragraph.
            builder.Writeln("Embedded Excel OLE object:");

            // Insert the Excel file as an embedded OLE object (not as an icon).
            // Parameters: fileName, isLinked (false = embed), asIcon (false = show content), presentation (null = default icon).
            builder.InsertOleObject(excelPath, false, false, null);

            // Build the output file name.
            string outputFileName = Path.GetFileNameWithoutExtension(wordFilePath) + "_WithOle.docx";
            string outputPath = Path.Combine(outputDir, outputFileName);

            // Save the modified document.
            doc.Save(outputPath);
        }

        Console.WriteLine("Processing completed.");
    }
}
