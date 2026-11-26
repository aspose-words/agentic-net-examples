using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PNG image.
        const string inputImagePath = "input.png";
        const int originalWidth = 200;
        const int originalHeight = 200;

        // Create bitmap using Aspose.Drawing.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(originalWidth, originalHeight))
        {
            // Draw deterministic content.
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black))
                {
                    graphics.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }

            // Save the sample image.
            bitmap.Save(inputImagePath);
        }

        // Verify the sample image was created.
        if (!File.Exists(inputImagePath))
            throw new Exception("Failed to create the sample PNG image.");

        // Step 2: Create a Word document and insert the PNG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Verify the document was saved.
        if (!File.Exists(docPath))
            throw new Exception("Failed to save the sample document.");

        // Step 3: Extract PNG images from the document, resize them to 75%, and save previews.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int previewIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image using Aspose.Drawing.
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    int previewWidth = (int)(originalBitmap.Width * 0.75);
                    int previewHeight = (int)(originalBitmap.Height * 0.75);

                    // Create a new bitmap for the resized preview.
                    using (Aspose.Drawing.Bitmap previewBitmap = new Aspose.Drawing.Bitmap(previewWidth, previewHeight))
                    {
                        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(previewBitmap))
                        {
                            graphics.Clear(Aspose.Drawing.Color.White);
                            graphics.DrawImage(originalBitmap, 0, 0, previewWidth, previewHeight);
                        }

                        string previewPath = $"preview-{previewIndex}.png";
                        previewBitmap.Save(previewPath);
                        previewIndex++;

                        // Validate that the preview file was created.
                        if (!File.Exists(previewPath))
                            throw new Exception($"Failed to save preview image '{previewPath}'.");
                    }
                }
            }
        }

        // Ensure at least one preview was generated.
        if (previewIndex == 0)
            throw new Exception("No images were extracted and resized from the document.");

        Console.WriteLine("Image preview generation completed successfully.");
    }
}
