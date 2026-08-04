using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a blank Word document.
        Document scannedDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(scannedDoc);

        // Generate an image with text using Aspose.Drawing.
        using (MemoryStream imageStream = new MemoryStream())
        {
            using (Bitmap bitmap = new Bitmap(400, 200))
            {
                // Create a graphics object from the bitmap.
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.White);

                    // Use a fully qualified Aspose.Drawing.Font.
                    using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                    {
                        using (SolidBrush brush = new SolidBrush(Color.Black))
                        {
                            graphics.DrawString("Sample OCR Text", font, brush, new PointF(10, 80));
                        }
                    }
                }

                // Save the bitmap to the memory stream as PNG.
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // Save the document as a regular PDF (non‑searchable).
        const string sourcePdfPath = "sample.pdf";
        scannedDoc.Save(sourcePdfPath, SaveFormat.Pdf);

        // Load the PDF and convert it to a searchable PDF/A‑1a document.
        Document pdfDoc = new Document(sourcePdfPath);
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Set compliance to PDF/A‑1a (searchable and tagged).
            Compliance = PdfCompliance.PdfA1a
        };

        const string outputPdfPath = "searchable_pdfa1a.pdf";
        pdfDoc.Save(outputPdfPath, saveOptions);

        // Verify that the output file was created.
        if (!File.Exists(outputPdfPath) || new FileInfo(outputPdfPath).Length == 0)
            throw new InvalidOperationException("The searchable PDF/A‑1a file was not created.");
    }
}
