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

        // Create a few sample DOCX files if the input folder is empty.
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"));
            CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"));
        }

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document sourceDoc = new Document(docxPath);

            // Ensure the layout is up‑to‑date and obtain the page count.
            int pageCount = sourceDoc.PageCount;

            // Extract each page and save it as an individual PDF.
            for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                // Extract a single page (pageIndex is zero‑based). The second parameter is the count of pages to extract.
                Document pageDoc = sourceDoc.ExtractPages(pageIndex, 1);

                // Build the output PDF file name.
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(docxPath)}_Page{pageIndex + 1}.pdf";
                string pdfPath = Path.Combine(outputDir, pdfFileName);

                // Save the extracted page as PDF.
                pageDoc.Save(pdfPath, SaveFormat.Pdf);
            }
        }

        // Validate that PDFs were created.
        int createdPdfCount = Directory.GetFiles(outputDir, "*.pdf").Length;
        if (createdPdfCount == 0)
        {
            throw new InvalidOperationException("No PDF files were generated.");
        }

        // Report the result.
        Console.WriteLine($"Processing complete. {createdPdfCount} PDF files created in '{outputDir}'.");
    }

    // Helper method to create a simple multi‑page DOCX document.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three pages with distinct content.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i} of {Path.GetFileName(filePath)}.");
            if (i < 3)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        // Save the document.
        doc.Save(filePath);
    }
}
