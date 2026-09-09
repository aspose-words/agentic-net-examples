using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output DOCX files.
        string inputFolder = "InputPdfs";
        string outputFolder = "OutputDocx";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample PDF files to work with.
        CreateSamplePdf(Path.Combine(inputFolder, "sample1.pdf"), "Sample PDF content 1");
        CreateSamplePdf(Path.Combine(inputFolder, "sample2.pdf"), "Sample PDF content 2");
        CreateSamplePdf(Path.Combine(inputFolder, "sample3.pdf"), "Sample PDF content 3");

        // Get all PDF files from the input folder.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");

        // Convert each PDF to DOCX and save to the output folder.
        foreach (string pdfPath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfPath);

            // Determine the output DOCX file path.
            string outputFileName = Path.GetFileNameWithoutExtension(pdfPath) + ".docx";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the document as DOCX.
            pdfDoc.Save(outputPath, SaveFormat.Docx);

            // Verify that the DOCX file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Conversion failed: '{outputPath}' was not created.");
        }

        // Example completed successfully.
        // (No interactive output required.)
    }

    // Helper method to create a simple PDF file with given text.
    private static void CreateSamplePdf(string filePath, string text)
    {
        // Create a blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);

        // Save the document as PDF.
        doc.Save(filePath, SaveFormat.Pdf);
    }
}
