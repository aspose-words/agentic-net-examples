using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files and output PDF files.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX documents if the input folder is empty.
        CreateSampleDocumentsIfNeeded(inputDir);

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document sourceDoc = new Document(docxPath);

            // Ensure the document has been laid out so PageCount is accurate.
            int pageCount = sourceDoc.PageCount;

            // Split the document page by page.
            for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                // Extract a single page (pageIndex is zero‑based, count is 1).
                Document pageDoc = sourceDoc.ExtractPages(pageIndex, 1);

                // Build the output PDF file name.
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(docxPath)}_Page{pageIndex + 1}.pdf";
                string pdfPath = Path.Combine(outputDir, pdfFileName);

                // Save the extracted page as PDF.
                pageDoc.Save(pdfPath, SaveFormat.Pdf);

                // Validate that the PDF was created.
                if (!File.Exists(pdfPath))
                {
                    throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
                }
            }
        }

        // Optional: write a simple completion message.
        Console.WriteLine("Document splitting completed.");
    }

    // Creates a couple of sample multi‑page DOCX files if none exist in the input folder.
    private static void CreateSampleDocumentsIfNeeded(string inputDir)
    {
        string[] existingDocs = Directory.GetFiles(inputDir, "*.docx");
        if (existingDocs.Length > 0)
            return; // Samples already exist.

        // Sample 1: three pages.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        for (int i = 1; i <= 3; i++)
        {
            builder1.Writeln($"Sample1 - Page {i}");
            if (i < 3)
                builder1.InsertBreak(BreakType.PageBreak);
        }
        string path1 = Path.Combine(inputDir, "Sample1.docx");
        doc1.Save(path1);

        // Sample 2: five pages.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        for (int i = 1; i <= 5; i++)
        {
            builder2.Writeln($"Sample2 - Page {i}");
            if (i < 5)
                builder2.InsertBreak(BreakType.PageBreak);
        }
        string path2 = Path.Combine(inputDir, "Sample2.docx");
        doc2.Save(path2);
    }
}
