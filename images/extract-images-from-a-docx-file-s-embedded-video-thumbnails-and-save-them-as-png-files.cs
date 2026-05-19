using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class ExtractVideoThumbnails
{
    public static void Main()
    {
        // Define file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "SampleWithVideo.docx");
        string thumbnailImagePath = Path.Combine(artifactsDir, "SampleThumbnail.png");

        // -------------------------------------------------
        // 1. Create a sample thumbnail image using Aspose.Drawing.
        // -------------------------------------------------
        int thumbWidth = 200;
        int thumbHeight = 150;
        using (Bitmap bitmap = new Bitmap(thumbWidth, thumbHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            // Draw a simple rectangle to make the thumbnail recognizable.
            graphics.DrawRectangle(new Pen(Aspose.Drawing.Color.DarkBlue, 3), 10, 10, thumbWidth - 20, thumbHeight - 20);
            bitmap.Save(thumbnailImagePath);
        }

        // -------------------------------------------------
        // 2. Create a DOCX document and insert the thumbnail image.
        //    In a real scenario this would be the video thumbnail.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image as an inline shape (simulating a video thumbnail).
        Shape thumbnailShape = builder.InsertImage(thumbnailImagePath);
        // Optionally set a name to identify it later.
        thumbnailShape.Name = "VideoThumbnail";

        // Save the document.
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all shape images.
        //    Save each extracted image as a PNG file.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Save the shape's image data to a memory stream.
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0; // Reset for reading.

                    // Load the image into Aspose.Drawing.Bitmap.
                    using (Bitmap bitmap = new Bitmap(imageStream))
                    {
                        // Ensure the output is PNG regardless of original format.
                        string outFile = Path.Combine(artifactsDir, $"Thumbnail_{imageIndex}.png");
                        bitmap.Save(outFile);
                        // Validate that the file was created.
                        if (!File.Exists(outFile))
                            throw new InvalidOperationException($"Failed to save extracted image: {outFile}");
                    }
                }
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // The program finishes here without awaiting user input.
    }
}
