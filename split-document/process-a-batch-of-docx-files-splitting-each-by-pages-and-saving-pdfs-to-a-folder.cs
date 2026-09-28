using System;
using System.IO;
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

        // Create sample DOCX files with multiple pages.
        CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"), "Sample Document 1", 3);
        CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"), "Sample Document 2", 4);

        // Process each DOCX file: split by pages and save each page as a PDF.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document sourceDoc = new Document(docxPath);
            int pageCount = sourceDoc.PageCount;

            // Aspose.Words uses zero‑based page indexes for extraction.
            for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                // Extract a single page (pageIndex is zero‑based, count = 1).
                Document pageDoc = sourceDoc.ExtractPages(pageIndex, 1);

                // Build the PDF file name and path (display pages as 1‑based).
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(docxPath)}_Page{pageIndex + 1}.pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the extracted page as PDF.
                pageDoc.Save(pdfPath, SaveFormat.Pdf);

                // Verify that the PDF was created.
                if (!File.Exists(pdfPath))
                {
                    throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
                }
            }
        }
    }

    // Helper method to create a sample DOCX with a given number of pages.
    private static void CreateSampleDocument(string filePath, string title, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln(title);
        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"Content for page {i}");
            if (i < pageCount)
            {
                // Insert a page break to start a new page.
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        doc.Save(filePath);
    }
}
