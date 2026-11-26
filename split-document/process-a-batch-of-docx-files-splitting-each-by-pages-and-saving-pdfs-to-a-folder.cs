using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files (2 documents, each with 3 pages).
        CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"));

        // Process each DOCX file: split by pages and save each page as a PDF.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document sourceDoc = new Document(docxPath);
            int pageCount = sourceDoc.PageCount;

            // Page indexes for ExtractPages are zero‑based, so iterate from 0.
            for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                // Extract a single page as a new document.
                Document pageDoc = sourceDoc.ExtractPages(pageIndex, 1);

                // Build output PDF file name (page numbers are 1‑based for readability).
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(docxPath)}_Page{pageIndex + 1}.pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the extracted page as PDF.
                pageDoc.Save(pdfPath);
            }
        }

        // Validation: ensure that a PDF was created for each page of each source document.
        int expectedPdfCount = 0;
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docxPath);
            expectedPdfCount += doc.PageCount;
        }

        int actualPdfCount = Directory.GetFiles(outputFolder, "*.pdf").Length;
        if (actualPdfCount != expectedPdfCount)
        {
            throw new InvalidOperationException(
                $"PDF split validation failed. Expected {expectedPdfCount} PDFs, but found {actualPdfCount}.");
        }
    }

    // Helper method to create a sample multi‑page DOCX document.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i} of the sample document.");
            // Insert a page break after each page except the last.
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(filePath);
    }
}
