using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Text;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Define temporary file paths.
        const string inputPdfPath = "input.pdf";
        const string outputPdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample image that contains non‑searchable text.
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(300, 100))
        {
            // Fill background with white.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);

                // Draw black text onto the bitmap.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                using (SolidBrush brush = new SolidBrush(Color.Black))
                {
                    graphics.DrawString("Sample scanned text", font, brush, new PointF(10, 30));
                }
            }

            // Save the bitmap to a memory stream as PNG.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // -----------------------------------------------------------------
                // 2. Insert the image into a blank Word document and save as PDF.
                // -----------------------------------------------------------------
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.InsertImage(imageStream);
                doc.Save(inputPdfPath, SaveFormat.Pdf);
            }
        }

        // -----------------------------------------------------------------
        // 3. Load the generated PDF.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(inputPdfPath);

        // -----------------------------------------------------------------
        // 4. Configure PDF/A‑2u compliance (PDF/A‑2b equivalent) and export structure.
        // -----------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,          // PDF/A‑2b equivalent.
            ExportDocumentStructure = true             // Required for PDF/A compliance.
        };

        // -----------------------------------------------------------------
        // 5. Save the PDF as a searchable PDF/A‑2u document.
        // -----------------------------------------------------------------
        pdfDoc.Save(outputPdfPath, saveOptions);

        // -----------------------------------------------------------------
        // 6. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The searchable PDF/A‑2u file was not created.");

        // Optional: clean up the intermediate file.
        if (File.Exists(inputPdfPath))
            File.Delete(inputPdfPath);
    }
}
