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
        // Deterministic file names.
        const string tiffImagePath = "sample.tif";
        const string docPath = "documentWithTiff.docx";

        // -------------------------------------------------
        // 1. Create a sample TIFF image using Aspose.Drawing.
        // -------------------------------------------------
        int width = 200;
        int height = 200;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple black rectangle.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 5))
                {
                    g.DrawRectangle(pen, 20, 20, width - 40, height - 40);
                }
            }

            // Save as TIFF.
            bitmap.Save(tiffImagePath, Aspose.Drawing.Imaging.ImageFormat.Tiff);
        }

        // Verify the TIFF image was created.
        if (!File.Exists(tiffImagePath))
            throw new FileNotFoundException("Failed to create the sample TIFF image.", tiffImagePath);

        // -------------------------------------------------
        // 2. Insert the TIFF image into a Word document.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(tiffImagePath);
        doc.Save(docPath);

        // Verify the document was saved.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Failed to save the document containing the TIFF image.", docPath);

        // -------------------------------------------------
        // 3. Load the document and extract images,
        //    converting each to a grayscale JPEG.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Set the image to display in grayscale.
            shape.ImageData.GrayScale = true;

            // Define output JPEG file name.
            string jpegFileName = $"extracted_{extractedCount}.jpg";

            // Save the image as JPEG. The file extension determines the format.
            shape.ImageData.Save(jpegFileName);

            // Validate that the JPEG file was created.
            if (!File.Exists(jpegFileName))
                throw new FileNotFoundException("Failed to save the grayscale JPEG image.", jpegFileName);

            extractedCount++;
        }

        // Ensure at least one image was extracted and converted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were found to extract and convert.");

        // -------------------------------------------------
        // 4. Clean up temporary files (optional).
        // -------------------------------------------------
        // File.Delete(tiffImagePath);
        // File.Delete(docPath);
    }
}
