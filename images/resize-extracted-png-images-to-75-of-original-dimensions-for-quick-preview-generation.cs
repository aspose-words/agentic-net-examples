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
        // Step 1: Create a deterministic sample PNG image.
        const int originalWidth = 200;
        const int originalHeight = 200;
        string inputImagePath = "input.png";

        Bitmap bitmap = new Bitmap(originalWidth, originalHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        Pen redPen = new Pen(Color.Red);
        graphics.DrawRectangle(redPen, 10, 10, originalWidth - 20, originalHeight - 20);
        bitmap.Save(inputImagePath, ImageFormat.Png);
        redPen.Dispose();
        graphics.Dispose();
        bitmap.Dispose();

        // Step 2: Insert the image into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = "document_with_image.docx";
        doc.Save(docPath);

        // Step 3: Extract the inserted PNG image from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        string extractedImagePath = "extracted.png";
        bool imageExtracted = false;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
            {
                shape.ImageData.Save(extractedImagePath);
                imageExtracted = true;
                break; // Only one image needed for this example.
            }
        }

        if (!imageExtracted || !File.Exists(extractedImagePath))
            throw new InvalidOperationException("Failed to extract the PNG image from the document.");

        // Step 4: Resize the extracted image to 75% of its original dimensions.
        using (Bitmap originalBitmap = new Bitmap(extractedImagePath))
        {
            int resizedWidth = (int)(originalBitmap.Width * 0.75);
            int resizedHeight = (int)(originalBitmap.Height * 0.75);

            using (Bitmap resizedBitmap = new Bitmap(resizedWidth, resizedHeight))
            {
                using (Graphics g = Graphics.FromImage(resizedBitmap))
                {
                    g.DrawImage(originalBitmap, 0, 0, resizedWidth, resizedHeight);
                }

                string previewImagePath = "preview.png";
                resizedBitmap.Save(previewImagePath, ImageFormat.Png);

                if (!File.Exists(previewImagePath))
                    throw new InvalidOperationException("Failed to save the resized preview image.");
            }
        }

        // All operations completed successfully.
    }
}
