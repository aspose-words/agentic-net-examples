using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string secureFolder = Path.Combine(artifactsDir, "SecureArchive");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(secureFolder);

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string pngPath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            // Draw a simple red ellipse.
            g.FillEllipse(Brushes.Red, 20, 20, 160, 160);
            // Save as PNG.
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Insert the PNG image into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);
        // Save the document (optional, just to have a file on disk).
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Extract PNG images from the document and convert them to
        //    grayscale BMP files saved in a secure folder.
        // -----------------------------------------------------------------
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                            .ToList();

        int imageIndex = 0;
        foreach (var shape in shapeNodes)
        {
            // Obtain the image bytes.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the PNG into a bitmap.
                using (Bitmap sourceBitmap = new Bitmap(imageStream))
                {
                    // Create a new bitmap for the grayscale version.
                    using (Bitmap grayBitmap = new Bitmap(sourceBitmap.Width, sourceBitmap.Height))
                    {
                        // Draw the source bitmap onto the new bitmap.
                        using (Graphics g = Graphics.FromImage(grayBitmap))
                        {
                            g.DrawImage(sourceBitmap, 0, 0, sourceBitmap.Width, sourceBitmap.Height);
                        }

                        // Convert each pixel to grayscale.
                        for (int y = 0; y < grayBitmap.Height; y++)
                        {
                            for (int x = 0; x < grayBitmap.Width; x++)
                            {
                                Color pixel = grayBitmap.GetPixel(x, y);
                                int gray = (int)(pixel.R * 0.3 + pixel.G * 0.59 + pixel.B * 0.11);
                                Color grayColor = Color.FromArgb(gray, gray, gray);
                                grayBitmap.SetPixel(x, y, grayColor);
                            }
                        }

                        // Save the grayscale bitmap as BMP in the secure folder.
                        string outputPath = Path.Combine(secureFolder, $"extracted_{imageIndex}.bmp");
                        grayBitmap.Save(outputPath, ImageFormat.Bmp);
                    }
                }
            }

            imageIndex++;
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one BMP file was created.
        // -----------------------------------------------------------------
        if (imageIndex == 0 || !Directory.EnumerateFiles(secureFolder, "*.bmp").Any())
        {
            throw new InvalidOperationException("No grayscale BMP files were created.");
        }

        // Program completed successfully.
    }
}
