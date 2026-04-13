using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Extract3DModelThumbnails
{
    public static void Main()
    {
        // Prepare deterministic file names.
        const string imagePath = "thumbnail.png";
        const string docPath = "sample.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample image that will act as a 3D model thumbnail.
        // -----------------------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle to make the image non‑empty.
                graphics.DrawRectangle(
                    new Pen(Aspose.Drawing.Color.Blue, 5),
                    20, 20, imgWidth - 40, imgHeight - 40);
            }

            // Save the bitmap to a file that will be inserted into the document.
            bitmap.Save(imagePath);
        }

        // -----------------------------------------------------------------
        // Step 2: Create a DOCX document and embed the image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image – Aspose.Words stores it inside a Shape.
        builder.InsertImage(imagePath);
        // Save the document to disk.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Step 3: Load the document and extract all embedded images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputFileName = $"extracted_{extractedCount}{extension}";

                // Save the image data to a PNG (or the native format) file.
                shape.ImageData.Save(outputFileName);
                extractedCount++;

                // Validate that the file was created.
                if (!File.Exists(outputFileName))
                    throw new InvalidOperationException($"Failed to create image file '{outputFileName}'.");
            }
        }

        // -----------------------------------------------------------------
        // Validation: ensure at least one image was extracted.
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Cleanup: optional removal of temporary files (comment out if inspection needed).
        // File.Delete(imagePath);
        // File.Delete(docPath);
    }
}
