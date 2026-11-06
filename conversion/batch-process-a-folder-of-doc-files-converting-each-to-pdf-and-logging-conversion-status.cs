using System;
using System.IO;
using Aspose.Words;

public class BatchDocToPdfConverter
{
    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = "InputDocs";
        string outputFolder = "OutputPdfs";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOC files in the input folder.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Sample{i}.doc");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document {i}.");
            sampleDoc.Save(docPath, SaveFormat.Doc);
        }

        // Process each DOC file in the input folder.
        string[] docFiles = Directory.GetFiles(inputFolder, "*.doc");
        foreach (string docFilePath in docFiles)
        {
            // Load the DOC file.
            Document doc = new Document(docFilePath);

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docFilePath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF for '{docFilePath}'.");

            // Log conversion status.
            Console.WriteLine($"Converted '{Path.GetFileName(docFilePath)}' to PDF successfully.");
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
