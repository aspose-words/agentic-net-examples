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
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample image using Aspose.Drawing (no System.Drawing usage).
        string imagePath = Path.Combine(outputDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);

                // Use an explicit Aspose.Drawing.Font declaration.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12);
                try
                {
                    graphics.DrawString(
                        "Test",
                        font,
                        Brushes.White,
                        new PointF(10, 40));
                }
                finally
                {
                    font.Dispose();
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a Word document, add styled text and the image.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Font.Name = "Arial";
        builder.Font.Size = 14;
        builder.Writeln("Sample MHTML content with style.");
        builder.InsertImage(imagePath);
        builder.Font.Bold = true;
        builder.Writeln("Bold text after image.");

        // Save the document as MHTML.
        string mhtmlPath = Path.Combine(outputDir, "sample.mhtml");
        sourceDoc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // Load the MHTML file.
        Document loadedDoc = new Document(mhtmlPath);

        // Convert the loaded document to PDF, preserving images and styles.
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file exists and contains data.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF conversion failed or resulted in an empty file.");

        // Optional: clean up temporary files (commented out to keep output for inspection).
        // File.Delete(imagePath);
        // File.Delete(mhtmlPath);
        // File.Delete(pdfPath);
    }
}
