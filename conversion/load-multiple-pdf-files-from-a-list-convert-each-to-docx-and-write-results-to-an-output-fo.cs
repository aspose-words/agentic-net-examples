using System;
using System.IO;
using Aspose.Words;

public class PdfToDocxBatchConverter
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = "input_pdfs";
        string outputDir = "output_docx";

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample PDF files to simulate existing inputs.
        CreateSamplePdf(Path.Combine(inputDir, "sample1.pdf"), "This is the first sample PDF.");
        CreateSamplePdf(Path.Combine(inputDir, "sample2.pdf"), "This is the second sample PDF.");
        CreateSamplePdf(Path.Combine(inputDir, "sample3.pdf"), "This is the third sample PDF.");

        // Retrieve all PDF files from the input folder.
        string[] pdfFiles = Directory.GetFiles(inputDir, "*.pdf");

        // Convert each PDF file to DOCX format.
        foreach (string pdfPath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfPath);

            // Determine the output DOCX file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfPath);
            string docxPath = Path.Combine(outputDir, fileNameWithoutExt + ".docx");

            // Save the document as DOCX.
            pdfDoc.Save(docxPath, SaveFormat.Docx);

            // Verify that the DOCX file was created.
            if (!File.Exists(docxPath))
            {
                throw new InvalidOperationException($"Conversion failed: {docxPath} was not created.");
            }
        }
    }

    private static void CreateSamplePdf(string filePath, string text)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);

        // Save the document as PDF.
        doc.Save(filePath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(filePath))
        {
            throw new InvalidOperationException($"Sample PDF creation failed: {filePath} was not created.");
        }
    }
}
