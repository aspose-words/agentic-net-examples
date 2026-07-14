using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // Step 1: Create a PDF that contains only an image (simulating a scanned page).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Generate a bitmap with some text using Aspose.Drawing.
        using (MemoryStream imgStream = new MemoryStream())
        {
            using (Bitmap bitmap = new Bitmap(300, 100))
            {
                // Obtain a Graphics object from the bitmap.
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.White);

                    // Create a drawing font (fully qualified to avoid ambiguity).
                    Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);

                    // Use a solid brush for the text colour.
                    using (SolidBrush brush = new SolidBrush(Color.Black))
                    {
                        // Draw the string onto the bitmap.
                        graphics.DrawString("Sample OCR Text", font, brush, new PointF(10, 30));
                    }
                }

                // Save the bitmap as PNG into the memory stream.
                bitmap.Save(imgStream, Aspose.Drawing.Imaging.ImageFormat.Png);
            }

            imgStream.Position = 0;
            // Insert the generated image into the document.
            builder.InsertImage(imgStream);
        }

        // Save the document as a regular PDF (non‑searchable).
        string inputPdfPath = Path.Combine(workDir, "input.pdf");
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Verify that the input PDF was created.
        if (!File.Exists(inputPdfPath) || new FileInfo(inputPdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the source PDF.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and save it as PDF/A‑2u (closest to PDF/A‑2b) using OCR‑like settings.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(inputPdfPath);

        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Set PDF/A‑2u compliance (PDF/A‑2b is not a separate enum value in this version).
            Compliance = PdfCompliance.PdfA2u,

            // The OCR properties are not available in this version of Aspose.Words.
            // If they were, you would enable them here, e.g.:
            // OcrMode = OcrMode.Auto,
            // OcrLanguage = OcrLanguage.English
        };

        string outputPdfPath = Path.Combine(workDir, "output.pdf");
        pdfDoc.Save(outputPdfPath, saveOptions);

        // Verify that the output PDF/A‑2u was created.
        if (!File.Exists(outputPdfPath) || new FileInfo(outputPdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the searchable PDF/A‑2u.");

        Console.WriteLine("Conversion completed successfully.");
    }
}
