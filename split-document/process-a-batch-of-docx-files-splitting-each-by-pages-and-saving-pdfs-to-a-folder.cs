using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentByPages
{
    public static void Main()
    {
        // Base working directory for the demo.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitDocumentDemo");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure input and output folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files with multiple pages.
        CreateSampleDocuments(inputDir);

        // Process each DOCX file: split by pages and save each page as a PDF.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document sourceDoc = new Document(docxPath);
            int pageCount = sourceDoc.PageCount; // Triggers layout calculation.

            for (int page = 1; page <= pageCount; page++)
            {
                // Extract a single page (zero‑based index, count = 1).
                Document pageDoc = sourceDoc.ExtractPages(page - 1, 1);

                // Build the output PDF file name.
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(docxPath)}_Page{page}.pdf";
                string pdfPath = Path.Combine(outputDir, pdfFileName);

                // Save the extracted page as PDF.
                pageDoc.Save(pdfPath, SaveFormat.Pdf);
            }
        }

        // Validate that each expected PDF file was created.
        ValidateOutputs(inputDir, outputDir);
    }

    private static void CreateSampleDocuments(string folder)
    {
        // Create two sample documents, each with three pages.
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int page = 1; page <= 3; page++)
            {
                builder.Writeln($"Sample Document {docIndex} - Page {page}");
                if (page < 3)
                {
                    // Insert a page break to start a new page.
                    builder.InsertBreak(BreakType.PageBreak);
                }
            }

            string fileName = Path.Combine(folder, $"Sample{docIndex}.docx");
            doc.Save(fileName);
        }
    }

    private static void ValidateOutputs(string inputFolder, string outputFolder)
    {
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docxPath);
            int expectedPages = doc.PageCount;

            for (int page = 1; page <= expectedPages; page++)
            {
                string expectedPdf = Path.Combine(
                    outputFolder,
                    $"{Path.GetFileNameWithoutExtension(docxPath)}_Page{page}.pdf");

                if (!File.Exists(expectedPdf))
                {
                    throw new FileNotFoundException($"Expected PDF not found: {expectedPdf}");
                }
            }
        }
    }
}
