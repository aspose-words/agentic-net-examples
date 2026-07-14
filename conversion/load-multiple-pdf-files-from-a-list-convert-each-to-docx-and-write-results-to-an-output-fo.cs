using System;
using System.IO;
using Aspose.Words;

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

        // Create sample PDF files to demonstrate the conversion.
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            // Create a new blank document and add some text.
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF content #{i}");

            // Save the document as PDF in the input folder.
            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Get all PDF files from the input folder.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");

        // Convert each PDF to DOCX and save to the output folder.
        foreach (string pdfFilePath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDocument = new Document(pdfFilePath);

            // Determine the output DOCX file path.
            string docxFileName = Path.GetFileNameWithoutExtension(pdfFilePath) + ".docx";
            string docxPath = Path.Combine(outputFolder, docxFileName);

            // Save the document as DOCX.
            pdfDocument.Save(docxPath, SaveFormat.Docx);

            // Verify that the DOCX file was created.
            if (!File.Exists(docxPath))
                throw new InvalidOperationException($"Failed to create output file: {docxPath}");

            Console.WriteLine($"Converted '{Path.GetFileName(pdfFilePath)}' to '{docxFileName}'.");
        }

        Console.WriteLine("Batch conversion completed successfully.");
    }
}
