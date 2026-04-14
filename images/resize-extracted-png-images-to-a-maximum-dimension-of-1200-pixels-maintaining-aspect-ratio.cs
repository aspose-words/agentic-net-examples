using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // ---------- Create a deterministic large PNG image ----------
        const string inputImagePath = "input.png";
        const int originalWidth = 2000;
        const int originalHeight = 1500;

        // Create bitmap and draw simple content.
        Bitmap sourceBitmap = new Bitmap(originalWidth, originalHeight);
        Graphics sourceGraphics = Graphics.FromImage(sourceBitmap);
        sourceGraphics.Clear(Color.LightBlue);
        sourceGraphics.FillRectangle(Brushes.DarkBlue, 100, 100, 1800, 1300);
        sourceGraphics.Dispose();

        // Save the source image to a local file.
        sourceBitmap.Save(inputImagePath);
        sourceBitmap.Dispose();

        // ---------- Insert the image into a Word document ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.InsertImage(inputImagePath);
        // InsertImage already appends the shape to the paragraph.
        doc.Save("DocumentWithImage.docx");

        // ---------- Load the document and resize PNG images ----------
        Document loadedDoc = new Document("DocumentWithImage.docx");
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape imgShape in shapeNodes.OfType<Shape>())
        {
            if (!imgShape.HasImage)
                continue;

            // Process only PNG images.
            if (imgShape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                imgShape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap originalBmp = new Bitmap(originalStream))
                {
                    int width = originalBmp.Width;
                    int height = originalBmp.Height;

                    // Determine scaling factor to limit max dimension to 1200px.
                    double scale = Math.Min(1200.0 / width, 1200.0 / height);
                    if (scale > 1.0) scale = 1.0; // Do not upscale.

                    int newWidth = (int)Math.Round(width * scale);
                    int newHeight = (int)Math.Round(height * scale);

                    // If resizing is needed, create a new bitmap and draw the scaled image.
                    using (Bitmap resizedBmp = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBmp))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(originalBmp, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized image to a deterministic file name.
                        string resizedPath = $"resized_{imageIndex}.png";
                        resizedBmp.Save(resizedPath);
                        if (!File.Exists(resizedPath))
                            throw new InvalidOperationException($"Failed to create resized image file: {resizedPath}");

                        // Replace the image in the document with the resized version via a stream.
                        using (MemoryStream resizedStream = new MemoryStream())
                        {
                            resizedBmp.Save(resizedStream, ImageFormat.Png);
                            resizedStream.Position = 0;
                            imgShape.ImageData.SetImage(resizedStream);
                        }
                    }
                }
            }

            imageIndex++;
        }

        // Save the document with resized images.
        loadedDoc.Save("DocumentWithResizedImages.docx");

        // Validate that at least one PNG image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found and resized in the document.");
    }
}
