using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfPagesToPng
{
    public static void Main()
    {
        // Create a sample multi‑page document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as a DOC file (bootstrap step).
        const string docPath = "input.doc";
        sourceDoc.Save(docPath, SaveFormat.Doc);

        // Load the DOC file and convert it to PDF.
        Document pdfDoc = new Document(docPath);
        const string pdfPath = "sample.pdf";
        pdfDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF for image export.
        Document loadedPdf = new Document(pdfPath);

        // Prepare image save options for PNG format.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        // Export only even‑numbered pages (pages 2,4,…) as separate PNG files.
        for (int pageIndex = 0; pageIndex < loadedPdf.PageCount; pageIndex++)
        {
            // Even‑numbered pages have odd zero‑based indices.
            if (pageIndex % 2 == 1)
            {
                pngOptions.PageSet = new PageSet(pageIndex);
                string pngPath = $"page_{pageIndex + 1}.png";
                loadedPdf.Save(pngPath, pngOptions);

                // Validate that the PNG file was created.
                if (!File.Exists(pngPath))
                    throw new InvalidOperationException($"PNG file '{pngPath}' was not created.");
            }
        }
    }
}
