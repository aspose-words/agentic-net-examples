using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output DOCX files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocx");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files.
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            Document pdfDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(pdfDoc);
            builder.Writeln($"This is the content of sample PDF #{i}.");
            pdfDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Gather all PDF files from the input folder.
        List<string> pdfFiles = new List<string>(Directory.GetFiles(inputFolder, "*.pdf"));

        // Convert each PDF to DOCX and save to the output folder.
        foreach (string pdfFilePath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDocument = new Document(pdfFilePath);

            // Determine the output DOCX file path.
            string docxFileName = Path.GetFileNameWithoutExtension(pdfFilePath) + ".docx";
            string docxPath = Path.Combine(outputFolder, docxFileName);

            // Save as DOCX.
            pdfDocument.Save(docxPath, SaveFormat.Docx);

            // Validate that the DOCX file was created.
            if (!File.Exists(docxPath))
                throw new InvalidOperationException($"Expected output DOCX was not created: {docxPath}");

            Console.WriteLine($"Converted '{Path.GetFileName(pdfFilePath)}' to '{docxFileName}'.");
        }

        Console.WriteLine("All PDF files have been successfully converted to DOCX.");
    }
}
