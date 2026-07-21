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
        // Create a deterministic sample JPEG image (2000x1500) to be inserted into the document.
        const int sampleWidth = 2000;
        const int sampleHeight = 1500;
        const string sampleImagePath = "sample.jpg";

        // Use Aspose.Drawing to create the bitmap and fill it with a solid background.
        using (Bitmap bitmap = new Bitmap(sampleWidth, sampleHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // Create a new document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "DocumentWithImage.docx";
        doc.Save(docPath);

        // Define the maximum dimensions for the resized images.
        const int maxWidth = 1024;
        const int maxHeight = 768;
        int imageIndex = 0;

        // Iterate over all Shape nodes that may contain images.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Retrieve the original image bytes and load them into an Aspose.Drawing.Bitmap.
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Bitmap originalBitmap = new Bitmap(ms))
            {
                int originalWidth = originalBitmap.Width;
                int originalHeight = originalBitmap.Height;

                // Compute scaling factor to fit within the target dimensions while preserving aspect ratio.
                double widthRatio = (double)maxWidth / originalWidth;
                double heightRatio = (double)maxHeight / originalHeight;
                double scale = Math.Min(widthRatio, heightRatio);
                if (scale > 1.0) // Do not upscale smaller images.
                    scale = 1.0;

                int newWidth = (int)Math.Round(originalWidth * scale);
                int newHeight = (int)Math.Round(originalHeight * scale);

                // Resize the image using Aspose.Drawing.
                using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                using (Graphics g = Graphics.FromImage(resizedBitmap))
                {
                    g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                    string resizedPath = $"resized_{imageIndex}.jpg";
                    resizedBitmap.Save(resizedPath, ImageFormat.Jpeg);
                }
            }

            imageIndex++;
        }

        // Validate that at least one JPEG image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to resize.");

        // Optional cleanup (commented out).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
