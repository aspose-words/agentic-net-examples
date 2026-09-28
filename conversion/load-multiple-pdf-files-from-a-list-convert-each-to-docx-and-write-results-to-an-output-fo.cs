using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class PdfToDocxBatchConverter
{
    public static void Main()
    {
        // Define folders for input PDFs and output DOCX files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocx");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files to simulate existing inputs.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF content {i}");
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Get all PDF files from the input folder.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");

        // Convert each PDF to DOCX.
        foreach (string pdfFilePath in pdfFiles)
        {
            // Load the PDF document using LoadOptions for PDF format.
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFormat = LoadFormat.Pdf
            };
            Document pdfDocument = new Document(pdfFilePath, loadOptions);

            // Determine the output DOCX file path.
            string outputFileName = Path.GetFileNameWithoutExtension(pdfFilePath) + ".docx";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the document as DOCX.
            pdfDocument.Save(outputPath, SaveFormat.Docx);

            // Validate that the DOCX file was created.
            if (!File.Exists(outputPath))
            {
                throw new InvalidOperationException($"Expected output DOCX was not created: {outputPath}");
            }
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("PDF to DOCX batch conversion completed successfully.");
    }
}
