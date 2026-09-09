using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputPdfs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Seed the input folder with a few sample DOC files.
        for (int i = 1; i <= 3; i++)
        {
            // Create a blank document and add sample content.
            Document source = new Document();
            DocumentBuilder builder = new DocumentBuilder(source);
            builder.Writeln($"Sample DOC content {i}.");

            // Save the document as a .doc file.
            string inputPath = Path.Combine(inputFolder, $"sample{i}.doc");
            source.Save(inputPath, SaveFormat.Doc);
        }

        // Process each .doc file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.doc"))
        {
            // Load the DOC file.
            Document doc = new Document(docPath);

            // Determine the corresponding PDF output path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Convert and save as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF for '{docPath}'.");

            // Log conversion status.
            Console.WriteLine($"Converted '{Path.GetFileName(docPath)}' to PDF successfully.");
        }

        // Indicate that batch processing is complete.
        Console.WriteLine("Batch conversion completed.");
    }
}
