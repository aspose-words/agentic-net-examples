using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page Word document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF – this PDF will be the source for conversion.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 2. Export only even‑numbered pages (2, 4, …) as separate PNG files.
        // -----------------------------------------------------------------
        const string outputBaseName = "even_pages.png";
        string baseFileNameWithoutExt = Path.GetFileNameWithoutExtension(outputBaseName);
        string extension = Path.GetExtension(outputBaseName);

        // Page indices are zero‑based; even‑numbered pages correspond to 1, 3, …
        List<string> createdFiles = new List<string>();
        for (int pageIndex = 1; pageIndex < pdfDoc.PageCount; pageIndex += 2)
        {
            // Prepare options for a single page.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(pageIndex, 1) // Export only this page.
            };

            // Determine the file name: first page uses the base name,
            // subsequent pages get a numeric suffix.
            string outputPath = createdFiles.Count == 0
                ? outputBaseName
                : $"{baseFileNameWithoutExt}_{createdFiles.Count}{extension}";

            pdfDoc.Save(outputPath, pngOptions);
            createdFiles.Add(outputPath);
        }

        // -----------------------------------------------------------------
        // 3. Validate that the expected PNG files were created.
        // -----------------------------------------------------------------
        foreach (string filePath in createdFiles)
        {
            if (!File.Exists(filePath))
                throw new InvalidOperationException($"Expected output file '{filePath}' was not created.");
        }

        // Optional: clean up generated files (comment out if you want to inspect them).
        //foreach (string filePath in createdFiles)
        //    File.Delete(filePath);
        //File.Delete(pdfPath);
    }
}
