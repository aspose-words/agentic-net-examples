using System;
using System.IO;
using System.Linq;
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

        // Create sample DOCX files if none exist.
        CreateSampleDocuments(inputFolder);

        // Process each DOCX file: split by pages and save each page as a PDF.
        var docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        int totalExpectedPdfs = 0;

        foreach (var docxPath in docxFiles)
        {
            Document doc = new Document(docxPath);
            int pageCount = doc.PageCount;
            totalExpectedPdfs += pageCount;

            for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                // Extract a single page (zero‑based index).
                Document pageDoc = doc.ExtractPages(pageIndex, 1);

                // Build output PDF file name.
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(docxPath)}_Page{pageIndex + 1}.pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the extracted page as PDF.
                pageDoc.Save(pdfPath, SaveFormat.Pdf);
            }
        }

        // Validate that the expected number of PDF files were created.
        int actualPdfCount = Directory.GetFiles(outputFolder, "*.pdf").Length;
        if (actualPdfCount != totalExpectedPdfs)
        {
            throw new InvalidOperationException(
                $"Expected {totalExpectedPdfs} PDF files, but found {actualPdfCount}.");
        }

        // Program completed successfully.
    }

    // Helper method to create a couple of sample multi‑page DOCX documents.
    private static void CreateSampleDocuments(string folder)
    {
        // Document 1: 3 pages of repeated text.
        string doc1Path = Path.Combine(folder, "Sample1.docx");
        if (!File.Exists(doc1Path))
        {
            Document doc1 = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc1);
            for (int i = 0; i < 150; i++) // Enough lines to span multiple pages.
            {
                builder.Writeln($"Sample document 1 - line {i + 1}");
            }
            doc1.Save(doc1Path);
        }

        // Document 2: 2 pages with headings.
        string doc2Path = Path.Combine(folder, "Sample2.docx");
        if (!File.Exists(doc2Path))
        {
            Document doc2 = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc2);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading for Document 2");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            for (int i = 0; i < 100; i++)
            {
                builder.Writeln($"Sample document 2 - paragraph {i + 1}");
            }
            doc2.Save(doc2Path);
        }
    }
}
