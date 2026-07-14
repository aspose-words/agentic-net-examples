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
        // Prepare working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample TIFF image using Aspose.Drawing.
        string tiffPath = Path.Combine(workDir, "sample.tif");
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 100))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                graphics.DrawString(
                    "Sample TIFF",
                    new Aspose.Drawing.Font("Arial", 12),
                    new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
                    new Aspose.Drawing.PointF(10, 40));
            }
            bitmap.Save(tiffPath, Aspose.Drawing.Imaging.ImageFormat.Tiff);
        }

        // 2. Insert the TIFF image into a Word document.
        string docPath = Path.Combine(workDir, "docWithTiff.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(tiffPath);
        doc.Save(docPath);

        // 3. Load the document and extract images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Convert the image to grayscale.
            shape.ImageData.GrayScale = true;

            // Save the grayscale image as JPEG.
            string jpegPath = Path.Combine(workDir, $"grayscale_{imageIndex}.jpg");
            shape.ImageData.Save(jpegPath);

            // Verify that the file was created.
            if (!File.Exists(jpegPath))
                throw new InvalidOperationException($"Failed to create JPEG file: {jpegPath}");

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found to convert.");

        // Optional cleanup (commented out).
        // File.Delete(tiffPath);
        // File.Delete(docPath);
    }
}
