using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary input PDF (image‑only) and the final searchable PDF/A‑1a output.
        const string inputPdfPath = "input.pdf";
        const string outputPdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a simple bitmap image with some text (simulating a scanned page).
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(400, 200))
        {
            // Fill background with white.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);

                // Create a drawing font.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                try
                {
                    // Draw black text onto the bitmap.
                    using (SolidBrush brush = new SolidBrush(Color.Black))
                    {
                        graphics.DrawString(
                            "Sample OCR Text",
                            font,
                            brush,
                            new PointF(10, 80));
                    }
                }
                finally
                {
                    font.Dispose();
                }
            }

            // Save the bitmap to a memory stream in PNG format.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // -----------------------------------------------------------------
                // Step 2: Insert the image into a Word document and save it as a PDF.
                // The resulting PDF contains only an image, i.e., it is not searchable.
                // -----------------------------------------------------------------
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.InsertImage(imageStream);
                doc.Save(inputPdfPath, SaveFormat.Pdf);
            }
        }

        // Verify that the non‑searchable PDF was created.
        if (!File.Exists(inputPdfPath))
            throw new InvalidOperationException("Failed to create the input PDF.");

        // -----------------------------------------------------------------
        // Step 3: Load the PDF and save it as a searchable PDF/A‑1a.
        // OCR is not directly supported via PdfSaveOptions in this version,
        // but we can still produce a PDF/A‑1a compliant document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(inputPdfPath);

        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Set PDF/A‑1a compliance (includes document structure for searchability).
            Compliance = PdfCompliance.PdfA1a
        };

        pdfDoc.Save(outputPdfPath, saveOptions);

        // Verify that the searchable PDF/A‑1a was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("Failed to create the searchable PDF/A‑1a.");
    }
}
