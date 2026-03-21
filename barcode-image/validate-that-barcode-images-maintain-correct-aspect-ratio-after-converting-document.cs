using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class BarcodeAspectRatioValidator
{
    // Tolerance for floating‑point comparison of aspect ratios.
    private const double AspectRatioTolerance = 0.001;

    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a temporary DOCX that contains a barcode‑like image.
        // -----------------------------------------------------------------
        string tempFolder = Path.Combine(Path.GetTempPath(), "BarcodeAspectRatioDemo");
        Directory.CreateDirectory(tempFolder);

        string docPath = Path.Combine(tempFolder, "Barcodes.docx");
        string pdfPath = Path.Combine(tempFolder, "Barcodes.pdf");

        CreateSampleDocumentWithImage(docPath);

        // -----------------------------------------------------------------
        // 2. Record the aspect ratios of all image shapes in the DOCX.
        // -----------------------------------------------------------------
        Document doc = new Document(docPath);
        List<double> docAspectRatios = GetImageAspectRatios(doc);

        // -----------------------------------------------------------------
        // 3. Save the document to PDF using PdfSaveOptions.
        // -----------------------------------------------------------------
        var pdfSaveOptions = new PdfSaveOptions
        {
            DownsampleOptions = new DownsampleOptions { DownsampleImages = false }
        };
        doc.Save(pdfPath, pdfSaveOptions);

        // -----------------------------------------------------------------
        // 4. Load the generated PDF back into a Document object.
        // -----------------------------------------------------------------
        var pdfLoadOptions = new PdfLoadOptions { SkipPdfImages = false };
        Document pdfDoc = new Document(pdfPath, pdfLoadOptions);

        // -----------------------------------------------------------------
        // 5. Record the aspect ratios of all image shapes in the PDF.
        // -----------------------------------------------------------------
        List<double> pdfAspectRatios = GetImageAspectRatios(pdfDoc);

        // -----------------------------------------------------------------
        // 6. Validate that the number of images matches and each aspect ratio
        //    is preserved within the defined tolerance.
        // -----------------------------------------------------------------
        if (docAspectRatios.Count != pdfAspectRatios.Count)
        {
            Console.WriteLine($"Image count mismatch: DOCX has {docAspectRatios.Count}, PDF has {pdfAspectRatios.Count}.");
            return;
        }

        for (int i = 0; i < docAspectRatios.Count; i++)
        {
            double docRatio = docAspectRatios[i];
            double pdfRatio = pdfAspectRatios[i];
            double diff = Math.Abs(docRatio - pdfRatio);

            if (diff > AspectRatioTolerance)
                Console.WriteLine($"Image {i + 1}: aspect ratio changed (DOCX={docRatio:F4}, PDF={pdfRatio:F4}).");
            else
                Console.WriteLine($"Image {i + 1}: aspect ratio preserved (ratio={docRatio:F4}).");
        }

        // Cleanup (optional)
        // Directory.Delete(tempFolder, true);
    }

    /// <summary>
    /// Creates a simple DOCX file containing a single PNG image.
    /// </summary>
    private static void CreateSampleDocumentWithImage(string filePath)
    {
        // 1x1 red pixel PNG (base64)
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        using (var ms = new MemoryStream(pngBytes))
        {
            // Insert the image; Aspose.Words will treat it as a Shape.
            builder.InsertImage(ms);
        }

        doc.Save(filePath);
    }

    /// <summary>
    /// Retrieves the aspect ratios (Width / Height) of all image shapes in the given document.
    /// </summary>
    private static List<double> GetImageAspectRatios(Document document)
    {
        var ratios = new List<double>();
        NodeCollection shapes = document.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.IsImage && shape.Height != 0)
                ratios.Add(shape.Width / shape.Height);
        }
        return ratios;
    }
}
