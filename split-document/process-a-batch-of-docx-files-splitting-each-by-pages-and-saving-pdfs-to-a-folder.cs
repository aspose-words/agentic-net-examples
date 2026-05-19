using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory of the application.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        // Input and output folders.
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files for the batch.
        CreateSampleDocuments(inputDir);

        // Process each DOCX file: split by pages and save each page as a PDF.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document sourceDoc = new Document(docxPath);
            int pageCount = sourceDoc.PageCount;

            // Extract each page individually.
            for (int pageNumber = 1; pageNumber <= pageCount; pageNumber++)
            {
                // ExtractPages uses zero‑based index and a count of pages to extract.
                Document pageDoc = sourceDoc.ExtractPages(pageNumber - 1, 1);

                // Build the output PDF file name.
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(docxPath)}_Page{pageNumber}.pdf";
                string pdfPath = Path.Combine(outputDir, pdfFileName);

                // Save the extracted page as PDF.
                pageDoc.Save(pdfPath, SaveFormat.Pdf);
            }
        }

        // Simple validation: ensure at least one PDF was created.
        string[] pdfFiles = Directory.GetFiles(outputDir, "*.pdf");
        if (pdfFiles.Length == 0)
            throw new InvalidOperationException("No PDF files were generated.");

        // List the generated PDFs.
        Console.WriteLine("Generated PDF files:");
        foreach (string pdf in pdfFiles)
            Console.WriteLine(pdf);
    }

    // Helper method to create a few sample DOCX files with multiple pages.
    private static void CreateSampleDocuments(string folder)
    {
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create three pages per document.
            for (int page = 1; page <= 3; page++)
            {
                builder.Writeln($"Sample Document {docIndex} - Page {page}");
                if (page < 3)
                    builder.InsertBreak(BreakType.PageBreak);
            }

            string filePath = Path.Combine(folder, $"Sample{docIndex}.docx");
            doc.Save(filePath);
        }
    }
}
