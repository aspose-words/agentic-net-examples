using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample PNG image.
        string inputImagePath = "input.png";
        CreateSamplePng(inputImagePath, 200, 200);

        // Create a Word document and insert the PNG image.
        string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Load the document (demonstrates load lifecycle).
        Document loadedDoc = new Document(docPath);

        // Extract PNG images and generate 50% size previews.
        int imageIndex = 0;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasImage)
            {
                // Save the extracted original image.
                string extractedPath = $"extracted-{imageIndex}.png";
                shape.ImageData.Save(extractedPath);

                // Resize the extracted image to 50% of its original dimensions.
                string previewPath = $"preview-{imageIndex}.png";
                ResizePngHalf(extractedPath, previewPath);

                // Validate that the preview file was created.
                if (!File.Exists(previewPath))
                    throw new InvalidOperationException($"Resized preview not created: {previewPath}");

                imageIndex++;
            }
        }

        // Ensure at least one preview image was generated.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string path, int width, int height)
    {
        Bitmap bitmap = new Bitmap(width, height);
        Graphics g = Graphics.FromImage(bitmap);
        g.Clear(Color.White);
        using (SolidBrush brush = new SolidBrush(Color.Blue))
        {
            g.FillRectangle(brush, 20, 20, width - 40, height - 40);
        }
        g.Dispose();
        bitmap.Save(path);
        bitmap.Dispose();
    }

    // Resizes a PNG image to 50% of its original size.
    private static void ResizePngHalf(string sourcePath, string destPath)
    {
        Bitmap sourceBitmap = new Bitmap(sourcePath);
        int newWidth = sourceBitmap.Width / 2;
        int newHeight = sourceBitmap.Height / 2;

        Bitmap destBitmap = new Bitmap(newWidth, newHeight);
        Graphics g = Graphics.FromImage(destBitmap);
        g.DrawImage(sourceBitmap, 0, 0, newWidth, newHeight);
        g.Dispose();

        destBitmap.Save(destPath);
        sourceBitmap.Dispose();
        destBitmap.Dispose();
    }
}
