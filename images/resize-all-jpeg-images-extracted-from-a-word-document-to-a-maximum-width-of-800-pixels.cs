using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleImagePath, 1200, 800); // 1200x800 pixels

        // 2. Build a Word document that contains the JPEG image multiple times.
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        CreateDocumentWithImages(inputDocPath, sampleImagePath);

        // 3. Load the document and process JPEG images.
        Document doc = new Document(inputDocPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int jpegCount = 0;
        int resizedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            jpegCount++;

            // Extract image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load image into Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap originalBitmap = new Bitmap(ms))
                {
                    int originalWidth = originalBitmap.Width;
                    int originalHeight = originalBitmap.Height;

                    // If width exceeds 800 pixels, resize while preserving aspect ratio.
                    if (originalWidth > 800)
                    {
                        double scale = 800.0 / originalWidth;
                        int newWidth = 800;
                        int newHeight = (int)Math.Round(originalHeight * scale);

                        using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                        {
                            using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                            {
                                graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                            }

                            // Save resized image to file.
                            string resizedImagePath = Path.Combine(artifactsDir, $"resized_{jpegCount}.jpg");
                            resizedBitmap.Save(resizedImagePath, ImageFormat.Jpeg);
                            resizedCount++;

                            // Replace the image in the document with the resized version.
                            using (MemoryStream resizedStream = new MemoryStream())
                            {
                                resizedBitmap.Save(resizedStream, ImageFormat.Jpeg);
                                resizedStream.Position = 0;
                                shape.ImageData.SetImage(resizedStream);
                            }
                        }
                    }
                }
            }
        }

        // Validate that at least one JPEG image was processed.
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG images were found in the document.");

        // Validate that at least one image was resized.
        if (resizedCount == 0)
            throw new InvalidOperationException("No JPEG images required resizing.");

        // Save the modified document.
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");
        doc.Save(outputDocPath);
    }

    // Creates a deterministic JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                // Draw a simple rectangle for visual distinction.
                graphics.FillRectangle(new SolidBrush(Color.Coral), width / 4, height / 4, width / 2, height / 2);
            }

            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Creates a Word document and inserts the specified image multiple times.
    private static void CreateDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image three times.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertParagraph();
            builder.InsertImage(imagePath);
        }

        doc.Save(docPath);
    }
}
