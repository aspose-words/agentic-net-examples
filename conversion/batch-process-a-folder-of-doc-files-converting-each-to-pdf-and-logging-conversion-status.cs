using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOC files and output PDF files.
        string inputFolder = "InputDocs";
        string outputFolder = "OutputPdfs";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOC files in the input folder.
        for (int i = 1; i <= 3; i++)
        {
            string docFileName = $"Sample{i}.doc";
            string docPath = Path.Combine(inputFolder, docFileName);

            // Create a new document and add sample content.
            Document source = new Document();
            DocumentBuilder builder = new DocumentBuilder(source);
            builder.Writeln($"This is the content of {docFileName}.");

            // Save the document as DOC.
            source.Save(docPath, SaveFormat.Doc);
        }

        // Process each DOC file: convert to PDF and log the status.
        string[] docFiles = Directory.GetFiles(inputFolder, "*.doc");
        foreach (string docPath in docFiles)
        {
            // Load the DOC file.
            Document doc = new Document(docPath);

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF for '{docPath}'.");

            // Log conversion status.
            Console.WriteLine($"Converted '{docPath}' to '{pdfPath}'.");
        }

        // Final summary.
        Console.WriteLine($"Conversion completed. PDFs are located in '{outputFolder}'.");
    }
}
