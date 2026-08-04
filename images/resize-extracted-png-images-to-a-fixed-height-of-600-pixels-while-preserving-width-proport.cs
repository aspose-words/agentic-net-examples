using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color

public class ResizeExtractedPngImages
{
    public static void Main()
    {
        // Define deterministic file names.
        const string inputImagePath = "input.png";
        const string documentPath = "DocumentWithImage.docx";

        // -------------------------------------------------
        // Step 1: Create a sample PNG image (800x400).
        // -------------------------------------------------
        int originalWidth = 800;
        int originalHeight = 400;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background with a solid color.
            graphics.Clear(Color.LightBlue);

            // Save the bitmap as a PNG file.
            bitmap.Save(inputImagePath);
        }

        // -------------------------------------------------
        // Step 2: Create a Word document and insert the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape imageShape = builder.InsertImage(inputImagePath);
        // Save the document (required by the lifecycle rule).
        doc.Save(documentPath);

        // -------------------------------------------------
        // Step 3: Load the document and extract PNG images.
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Retrieve the image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // Ensure the stream is at the beginning.
                using (Bitmap originalBitmap = new Bitmap(ms))
                {
                    // -------------------------------------------------
                    // Step 4: Compute new dimensions (height = 600px, preserve aspect ratio).
                    // -------------------------------------------------
                    const int targetHeight = 600;
                    int originalImgHeight = originalBitmap.Height;
                    int originalImgWidth = originalBitmap.Width;

                    // Guard against zero height to avoid division by zero.
                    if (originalImgHeight == 0)
                        throw new InvalidOperationException("Original image height is zero.");

                    double scaleFactor = (double)targetHeight / originalImgHeight;
                    int targetWidth = (int)Math.Round(originalImgWidth * scaleFactor);

                    // -------------------------------------------------
                    // Step 5: Create a new bitmap with the target size and draw the scaled image.
                    // -------------------------------------------------
                    using (Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight))
                    using (Graphics g = Graphics.FromImage(resizedBitmap))
                    {
                        // Draw the original bitmap onto the new bitmap, scaling it.
                        g.DrawImage(originalBitmap, 0, 0, targetWidth, targetHeight);

                        // -------------------------------------------------
                        // Step 6: Save the resized image to a deterministic file name.
                        // -------------------------------------------------
                        string resizedImagePath = $"ResizedImage_{imageIndex}.png";
                        resizedBitmap.Save(resizedImagePath);

                        // Validate that the file was created.
                        if (!File.Exists(resizedImagePath))
                            throw new FileNotFoundException($"Failed to create resized image file: {resizedImagePath}");

                        Console.WriteLine($"Resized image saved: {resizedImagePath} (Width={targetWidth}, Height={targetHeight})");
                    }
                }
            }

            imageIndex++;
        }

        // If no PNG images were processed, indicate the situation.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found to resize.");
    }
}
