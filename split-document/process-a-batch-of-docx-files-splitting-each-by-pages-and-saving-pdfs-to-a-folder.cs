using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files if the input folder is empty.
        if (!Directory.GetFiles(inputFolder, "*.docx").Any())
        {
            CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"), 3);
            CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"), 5);
        }

        // Process each DOCX file: split by pages and save each page as a separate PDF.
        var docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        int totalExpectedPdfs = 0;

        foreach (var docxPath in docxFiles)
        {
            // Load the source document.
            Document sourceDoc = new Document(docxPath);

            // Ensure layout is up‑to‑date so PageCount is accurate.
            sourceDoc.UpdatePageLayout();

            int pageCount = sourceDoc.PageCount;
            totalExpectedPdfs += pageCount;

            // Extract each page (zero‑based index) and save it as an individual PDF.
            for (int pageIndex = 1; pageIndex <= pageCount; pageIndex++)
            {
                // Extract a single page. ExtractPages expects a zero‑based start index.
                Document pageDoc = sourceDoc.ExtractPages(pageIndex - 1, 1);

                // Build the output PDF file name.
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(docxPath)}_Page_{pageIndex}.pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the extracted page as PDF.
                pageDoc.Save(pdfPath, SaveFormat.Pdf);
            }
        }

        // Validation: ensure the expected number of PDF files were created.
        int actualPdfCount = Directory.GetFiles(outputFolder, "*.pdf").Length;
        if (actualPdfCount != totalExpectedPdfs)
        {
            throw new InvalidOperationException(
                $"PDF split validation failed. Expected {totalExpectedPdfs} PDFs but found {actualPdfCount}.");
        }

        // Indicate successful completion.
        Console.WriteLine("Document splitting completed successfully.");
    }

    // Helper method to create a simple multi‑page DOCX document.
    private static void CreateSampleDocument(string filePath, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"This is page {i} of {Path.GetFileNameWithoutExtension(filePath)}.");
            if (i < pageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(filePath, SaveFormat.Docx);
    }
}
