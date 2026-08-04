using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = "InputDocs";
        string outputFolder = "OutputPdfs";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOC files using the provided creation pattern.
        for (int i = 1; i <= 3; i++)
        {
            // Create a blank document.
            Document source = new Document();

            // Add sample content.
            DocumentBuilder builder = new DocumentBuilder(source);
            builder.Writeln($"Sample content for document {i}.");

            // Save as DOC.
            string inputPath = Path.Combine(inputFolder, $"Sample{i}.doc");
            source.Save(inputPath, SaveFormat.Doc);
        }

        // Batch process each DOC file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.doc"))
        {
            // Determine the corresponding PDF output path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docPath);
            string pdfPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.pdf");

            // Load the DOC file.
            Document doc = new Document(docPath);

            // Convert and save to PDF using the provided conversion pattern.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (File.Exists(pdfPath))
            {
                Console.WriteLine($"Converted '{docPath}' to '{pdfPath}'.");
            }
            else
            {
                throw new InvalidOperationException($"Failed to create PDF for '{docPath}'.");
            }
        }

        // Indicate completion.
        Console.WriteLine("Batch conversion completed.");
    }
}
