using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file and folder names.
        string inputImagePath = "input.png";
        string documentPath = "document.docx";
        string previewFolder = "previews";

        // Ensure the preview folder exists.
        Directory.CreateDirectory(previewFolder);

        // -------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -------------------------------------------------
        int originalWidth = 200;
        int originalHeight = 200;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Color.White);
                // Draw a simple blue rectangle.
                using (Brush brush = new SolidBrush(Color.Blue))
                {
                    g.FillRectangle(brush, 0, 0, originalWidth, originalHeight);
                }
            }
            // Save the image to a deterministic file name.
            bitmap.Save(inputImagePath);
        }

        // -------------------------------------------------
        // 2. Insert the image into a Word document.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(documentPath);

        // -------------------------------------------------
        // 3. Load the document and extract PNG images.
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .Cast<Shape>()
                                  .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                                  .ToList();

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Save the image data to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the original image using Aspose.Drawing.
                using (Image originalImage = Image.FromStream(imageStream))
                {
                    // Calculate 50% dimensions.
                    int newWidth = originalImage.Width / 2;
                    int newHeight = originalImage.Height / 2;

                    // Create a new bitmap for the resized preview.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.Clear(Color.White);
                            // Draw the original image scaled down.
                            graphics.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized preview.
                        string previewPath = Path.Combine(previewFolder, $"preview_{imageIndex}.png");
                        resizedBitmap.Save(previewPath);
                    }
                }
            }

            imageIndex++;
        }

        // -------------------------------------------------
        // 4. Validation: ensure at least one preview was created.
        // -------------------------------------------------
        int previewCount = Directory.GetFiles(previewFolder, "*.png").Length;
        if (previewCount == 0)
        {
            throw new InvalidOperationException("No preview images were generated.");
        }

        // Optional: clean up intermediate files (commented out to keep artifacts).
        // File.Delete(inputImagePath);
        // File.Delete(documentPath);
    }
}
