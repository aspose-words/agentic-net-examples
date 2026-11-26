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
        // Create a deterministic sample thumbnail image (PNG) using Aspose.Drawing.
        const string thumbnailPath = "thumb.png";
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Ensure any previous file is removed.
        if (File.Exists(thumbnailPath))
            File.Delete(thumbnailPath);

        // Create bitmap, draw simple rectangle, and save.
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.DrawRectangle(Pens.Black, 10, 10, imgWidth - 20, imgHeight - 20);
            }
            bitmap.Save(thumbnailPath, ImageFormat.Png);
        }

        // Verify that the thumbnail image was created.
        if (!File.Exists(thumbnailPath))
            throw new InvalidOperationException($"Failed to create thumbnail image at '{thumbnailPath}'.");

        // Create a DOCX document and insert the thumbnail image.
        const string docPath = "sample.docx";

        // Remove any existing document file.
        if (File.Exists(docPath))
            File.Delete(docPath);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(thumbnailPath);
        doc.Save(docPath);

        // Verify that the document was saved.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to save document at '{docPath}'.");

        // Load the document and extract all embedded images (including 3D model thumbnails).
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extractedFileName = $"extracted-{extractedCount}.png";

                // Remove any existing file with the same name.
                if (File.Exists(extractedFileName))
                    File.Delete(extractedFileName);

                shape.ImageData.Save(extractedFileName);
                // Validate that the image file was created.
                if (!File.Exists(extractedFileName))
                    throw new InvalidOperationException($"Failed to save extracted image '{extractedFileName}'.");
                extractedCount++;
            }
        }

        // Validation: ensure at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        Console.WriteLine($"Successfully extracted {extractedCount} image(s).");
        // Optional cleanup (uncomment if you want to delete temporary files).
        // File.Delete(thumbnailPath);
        // File.Delete(docPath);
    }
}
