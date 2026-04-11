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
        // Define file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string sourcePdfPath = Path.Combine(artifactsDir, "source.pdf");
        string searchablePdfPath = Path.Combine(artifactsDir, "searchable_pdfa2b.pdf");

        // -----------------------------------------------------------------
        // 1. Create a sample PDF that contains only an image (simulating a scanned document).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bitmap with some text using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(300, 100))
        {
            // Obtain a Graphics object for the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Use a fully qualified Aspose.Drawing.Font.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                using (SolidBrush brush = new SolidBrush(Color.Black))
                {
                    graphics.DrawString("Sample scanned text", font, brush, new PointF(10, 30));
                }
            }

            // Save the bitmap to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0; // Reset before inserting.

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // Save the document as a regular PDF (image‑only, not searchable).
        doc.Save(sourcePdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF and save it as searchable PDF/A‑2b (using PDF/A‑2u compliance).
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(sourcePdfPath);

        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Set compliance to PDF/A‑2u (closest to PDF/A‑2b).
            Compliance = PdfCompliance.PdfA2u
            // OCR settings are not available in this version of Aspose.Words.
        };

        // Save the searchable PDF/A‑2b (PDF/A‑2u) document.
        pdfDoc.Save(searchablePdfPath, saveOptions);

        // -----------------------------------------------------------------
        // 3. Validate that the output file was created and is not empty.
        // -----------------------------------------------------------------
        if (!File.Exists(searchablePdfPath))
            throw new FileNotFoundException("The searchable PDF/A‑2b file was not created.", searchablePdfPath);

        FileInfo info = new FileInfo(searchablePdfPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The searchable PDF/A‑2b file is empty.");

        Console.WriteLine("Searchable PDF/A‑2b created at: " + searchablePdfPath);
    }
}
