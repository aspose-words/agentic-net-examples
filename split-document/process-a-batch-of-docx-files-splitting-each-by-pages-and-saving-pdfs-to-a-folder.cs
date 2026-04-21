using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure input and output folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files.
        CreateSampleDocument(Path.Combine(inputDir, "SampleDoc1.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "SampleDoc2.docx"));

        // Keep track of expected PDF counts for validation.
        var expectedPdfCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document sourceDoc = new Document(docxPath);

            // Ensure layout is up‑to‑date so PageCount is accurate.
            sourceDoc.UpdatePageLayout();

            // Determine the number of pages.
            int pageCount = sourceDoc.PageCount;

            // Record expected number of PDFs for this document.
            expectedPdfCounts[docxPath] = pageCount;

            // Split the document page by page.
            for (int pageIndex = 1; pageIndex <= pageCount; pageIndex++)
            {
                // Extract a single page (zero‑based index, count = 1).
                Document pageDoc = sourceDoc.ExtractPages(pageIndex - 1, 1);

                // Build the output PDF file name.
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(docxPath)}_Page{pageIndex}.pdf";
                string pdfPath = Path.Combine(outputDir, pdfFileName);

                // Save the extracted page as PDF.
                pageDoc.Save(pdfPath, SaveFormat.Pdf);
            }
        }

        // Validate that the expected PDF files were created.
        foreach (var kvp in expectedPdfCounts)
        {
            string sourcePath = kvp.Key;
            int expectedCount = kvp.Value;
            string baseName = Path.GetFileNameWithoutExtension(sourcePath);

            for (int pageIndex = 1; pageIndex <= expectedCount; pageIndex++)
            {
                string expectedPdf = Path.Combine(outputDir, $"{baseName}_Page{pageIndex}.pdf");
                if (!File.Exists(expectedPdf))
                {
                    throw new Exception($"Expected PDF not found: {expectedPdf}");
                }
            }
        }

        // All PDFs generated successfully.
        Console.WriteLine("Document splitting completed successfully.");
    }

    // Helper method to create a sample multi‑page DOCX document.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three pages with distinct content.
        builder.Writeln($"This is page 1 of {Path.GetFileNameWithoutExtension(filePath)}.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln($"This is page 2 of {Path.GetFileNameWithoutExtension(filePath)}.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln($"This is page 3 of {Path.GetFileNameWithoutExtension(filePath)}.");

        // Save the document.
        doc.Save(filePath, SaveFormat.Docx);
    }
}
