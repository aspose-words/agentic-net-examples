using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string sampleImagePath = "sample.jpg";
        const string originalDocPath = "original.docx";
        const string resizedDocPath = "resized.docx";

        // -------------------------------------------------
        // 1. Create a sample JPEG image (1200x800 pixels)
        // -------------------------------------------------
        Bitmap bitmap = new Bitmap(1200, 800);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a simple rectangle to make the image non‑empty
        graphics.FillRectangle(new SolidBrush(Color.LightBlue), 100, 100, 1000, 600);
        graphics.Dispose();
        bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Build a Word document that contains the JPEG image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.InsertBreak(BreakType.PageBreak);
        builder.InsertImage(sampleImagePath);
        doc.Save(originalDocPath);

        // -------------------------------------------------
        // 3. Load the document and resize JPEG images to max 800px width
        // -------------------------------------------------
        Document loadedDoc = new Document(originalDocPath);
        var jpegShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .Cast<Shape>()
                                  .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                                  .ToList();

        if (!jpegShapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the document.");

        foreach (Shape shape in jpegShapes)
        {
            // Extract the image into a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading

                // Load the image with Aspose.Drawing
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    int originalWidth = originalBitmap.Width;
                    if (originalWidth <= 800)
                        continue; // No resizing needed

                    // Calculate new dimensions while preserving aspect ratio
                    double scale = 800.0 / originalWidth;
                    int newWidth = 800;
                    int newHeight = (int)(originalBitmap.Height * scale);

                    // Create a new bitmap with the target size
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    using (Graphics g = Graphics.FromImage(resizedBitmap))
                    {
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        g.Dispose();

                        // Save resized image to a new stream
                        using (MemoryStream resizedStream = new MemoryStream())
                        {
                            resizedBitmap.Save(resizedStream, ImageFormat.Jpeg);
                            resizedStream.Position = 0; // Reset before setting

                            // Replace the shape's image with the resized one
                            shape.ImageData.SetImage(resizedStream);
                        }
                    }
                }
            }
        }

        // -------------------------------------------------
        // 4. Save the modified document
        // -------------------------------------------------
        loadedDoc.Save(resizedDocPath);

        // -------------------------------------------------
        // 5. Validate that the output file was created
        // -------------------------------------------------
        if (!File.Exists(resizedDocPath))
            throw new FileNotFoundException("The resized document was not saved.", resizedDocPath);
    }
}
