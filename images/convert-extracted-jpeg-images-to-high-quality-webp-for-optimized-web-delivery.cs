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
        // Folder for all generated files.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic JPEG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // White background.
                g.Clear(Color.White);

                // Draw a red circle.
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }

            // Save as JPEG.
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Insert the JPEG into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        string docPath = Path.Combine(artifactsDir, "doc_with_image.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Convert the document page (which contains the JPEG) to WebP.
        // -----------------------------------------------------------------
        string webpPath = Path.Combine(artifactsDir, "converted_page.webp");
        ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP);
        // High‑quality WebP: keep default settings (Aspose.Words uses lossless for WebP when possible).
        doc.Save(webpPath, webpOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the WebP file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(webpPath))
            throw new Exception("WebP conversion failed.");

        // Example completed successfully.
    }
}
