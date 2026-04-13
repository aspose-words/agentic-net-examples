using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample image (simulating a QR code) and save it locally.
        // -----------------------------------------------------------------
        const string sampleImagePath = "sample.png";
        // Create a deterministic 200x200 white bitmap.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // Draw a simple black rectangle to act as placeholder QR content.
        graphics.FillRectangle(Aspose.Drawing.Brushes.Black, 50, 50, 100, 100);
        // Save the bitmap.
        bitmap.Save(sampleImagePath);
        // Clean up drawing objects.
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Build a Word document and insert the sample image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image file; this creates a Shape with an image.
        builder.InsertImage(sampleImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document, extract all images, and save them.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Ensure the shape actually contains image data.
            if (!shape.HasImage) continue;

            // Determine appropriate file extension based on the image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string extractedImagePath = $"extracted_{extractedCount}{extension}";

            // Save the image to the file system.
            shape.ImageData.Save(extractedImagePath);

            // Validate that the file was created.
            if (!File.Exists(extractedImagePath))
                throw new Exception($"Failed to save extracted image: {extractedImagePath}");

            Console.WriteLine($"Image #{extractedCount} saved as: {extractedImagePath}");
            extractedCount++;
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one image was extracted.
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 5. Optional cleanup of temporary files.
        // -----------------------------------------------------------------
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
        // foreach (var file in Directory.GetFiles(Directory.GetCurrentDirectory(), "extracted_*"))
        //     File.Delete(file);
    }
}
