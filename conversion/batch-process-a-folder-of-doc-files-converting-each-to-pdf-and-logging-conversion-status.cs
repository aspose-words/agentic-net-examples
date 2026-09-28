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
            // Create a new empty document.
            Document source = new Document();
            DocumentBuilder builder = new DocumentBuilder(source);
            builder.Writeln($"Sample content for document {i}.");

            // Save the document as DOC.
            string docPath = Path.Combine(inputFolder, $"Sample{i}.doc");
            source.Save(docPath, SaveFormat.Doc);
        }

        // Process each DOC file in the input folder.
        foreach (string docFilePath in Directory.GetFiles(inputFolder, "*.doc"))
        {
            // Load the DOC file.
            Document doc = new Document(docFilePath);

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docFilePath) + ".pdf";
            string pdfFilePath = Path.Combine(outputFolder, pdfFileName);

            // Convert and save as PDF.
            doc.Save(pdfFilePath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfFilePath))
                throw new InvalidOperationException($"Expected PDF was not created for '{docFilePath}'.");

            // Log conversion status.
            Console.WriteLine($"Converted '{docFilePath}' to '{pdfFilePath}'.");
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch conversion completed.");
    }
}
