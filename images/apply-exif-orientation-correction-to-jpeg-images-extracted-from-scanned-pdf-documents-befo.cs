using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExifOrientationCorrection
{
    public static void Main()
    {
        // Prepare output folder
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic JPEG image
        string sampleJpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleJpegPath);

        // 2. Insert the JPEG into a Word document and save as PDF (simulating a scanned PDF)
        string sourcePdfPath = Path.Combine(artifactsDir, "source.pdf");
        CreatePdfWithImage(sampleJpegPath, sourcePdfPath);

        // 3. Load the PDF, rotate each embedded image 90° clockwise (EXIF‑orientation simulation) and save
        string correctedPdfPath = Path.Combine(artifactsDir, "corrected.pdf");
        ApplyExifCorrectionAndSave(sourcePdfPath, correctedPdfPath);

        // 4. Verify output
        if (!File.Exists(correctedPdfPath))
            throw new InvalidOperationException("Corrected PDF was not created.");

        Console.WriteLine("EXIF orientation correction completed successfully.");
    }

    // -------------------------------------------------------------------------
    // Creates a simple JPEG image using Aspose.Drawing.
    // -------------------------------------------------------------------------
    private static void CreateSampleJpeg(string filePath)
    {
        const int width = 200;
        const int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, width - 40, height - 40);
                }
            }

            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // -------------------------------------------------------------------------
    // Inserts the image into a Word document and saves it as PDF.
    // -------------------------------------------------------------------------
    private static void CreatePdfWithImage(string imagePath, string pdfPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);
    }

    // -------------------------------------------------------------------------
    // Loads the PDF, rotates every image 90° clockwise and saves the result.
    // -------------------------------------------------------------------------
    private static void ApplyExifCorrectionAndSave(string inputPdf, string outputPdf)
    {
        // Load the PDF (Aspose.Words can load PDF directly)
        Document doc = new Document(inputPdf);

        // Find all shapes that contain an image
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int processedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue; // skip shapes without image data

            // Extract the original image into a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // reset before reading

                // Load the image with Aspose.Drawing
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Rotate 90° clockwise
                    using (Bitmap rotatedBitmap = RotateBitmap90Clockwise(originalBitmap))
                    {
                        // Save rotated bitmap into a new memory stream using the original format
                        using (MemoryStream rotatedStream = new MemoryStream())
                        {
                            ImageFormat targetFormat = GetImageFormat(shape.ImageData.ImageType);
                            rotatedBitmap.Save(rotatedStream, targetFormat);
                            rotatedStream.Position = 0; // reset for SetImage

                            // Replace the shape's image with the rotated one
                            shape.ImageData.SetImage(rotatedStream);
                        }
                    }
                }
            }

            processedCount++;
        }

        if (processedCount == 0)
            throw new InvalidOperationException("No images were found in the PDF.");

        // Save the corrected document as PDF
        doc.Save(outputPdf, SaveFormat.Pdf);
    }

    // -------------------------------------------------------------------------
    // Rotates a bitmap 90° clockwise.
    // -------------------------------------------------------------------------
    private static Bitmap RotateBitmap90Clockwise(Bitmap source)
    {
        int srcWidth = source.Width;
        int srcHeight = source.Height;

        // Width/height are swapped for a 90° rotation
        Bitmap rotated = new Bitmap(srcHeight, srcWidth);
        using (Graphics g = Graphics.FromImage(rotated))
        {
            // Move origin to centre of the new bitmap
            g.TranslateTransform(srcHeight / 2f, srcWidth / 2f);
            // Rotate 90° clockwise
            g.RotateTransform(90);
            // Move origin back and draw the original image
            g.TranslateTransform(-srcWidth / 2f, -srcHeight / 2f);
            g.DrawImage(source, 0, 0, srcWidth, srcHeight);
        }

        return rotated;
    }

    // -------------------------------------------------------------------------
    // Maps Aspose.Words.ImageType to Aspose.Drawing.Imaging.ImageFormat.
    // -------------------------------------------------------------------------
    private static ImageFormat GetImageFormat(ImageType imageType)
    {
        // Only map formats that are known to exist in Aspose.Drawing.Imaging.ImageFormat.
        return imageType switch
        {
            ImageType.Jpeg => ImageFormat.Jpeg,
            ImageType.Png => ImageFormat.Png,
            ImageType.Bmp => ImageFormat.Bmp,
            ImageType.Gif => ImageFormat.Gif,
            ImageType.Emf => ImageFormat.Emf,
            ImageType.Wmf => ImageFormat.Wmf,
            // For any other or unknown types, fall back to PNG which is widely supported.
            _ => ImageFormat.Png
        };
    }
}
