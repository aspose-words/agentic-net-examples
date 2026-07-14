using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;               // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;      // For ImageFormat and RotateFlipType

public class ExifOrientationCorrection
{
    public static void Main()
    {
        // Directories for temporary files
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample JPEG image (simulating a scanned page with wrong orientation)
        string originalJpegPath = Path.Combine(workDir, "scanned.jpg");
        CreateSampleJpeg(originalJpegPath);

        // 2. Insert the image into a Word document and save as PDF (simulating a scanned PDF)
        string pdfPath = Path.Combine(workDir, "ScannedDocument.pdf");
        CreatePdfWithImage(originalJpegPath, pdfPath);

        // 3. Load the PDF and extract JPEG images
        Document pdfDoc = new Document(pdfPath);
        NodeCollection shapes = pdfDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // 4. Save the original image to a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // 5. Load the image with Aspose.Drawing
                using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    // 6. Apply EXIF orientation correction.
                    //    For demonstration we assume the image needs a 90° clockwise rotation.
                    bitmap.RotateFlip(RotateFlipType.Rotate90FlipNone);

                    // 7. Save the corrected image to a deterministic file name
                    string correctedPath = Path.Combine(workDir,
                        $"CorrectedImage.{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");
                    bitmap.Save(correctedPath, ImageFormat.Jpeg);

                    // 8. Validate that the file was created
                    if (!File.Exists(correctedPath))
                        throw new InvalidOperationException($"Failed to save corrected image: {correctedPath}");

                    Console.WriteLine($"Corrected image saved: {correctedPath}");
                }
            }

            imageIndex++;
        }

        // Clean up (optional)
        // Directory.Delete(workDir, true);
    }

    // Creates a simple JPEG image with some text – this will act as the scanned page.
    private static void CreateSampleJpeg(string filePath)
    {
        int width = 400;
        int height = 300;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.White);
            g.DrawString(
                "Scanned Page",
                new Aspose.Drawing.Font("Arial", 24),
                new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
                new Aspose.Drawing.PointF(50, 120));

            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Inserts the JPEG into a Word document and saves it as PDF.
    private static void CreatePdfWithImage(string imagePath, string pdfPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
