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
        // Folder for all generated files.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample preview image that will represent an audio track.
        // -----------------------------------------------------------------
        string previewPath = Path.Combine(outputDir, "audio_preview.png");
        int previewWidth = 200;
        int previewHeight = 200;

        // Use Aspose.Drawing to create a deterministic bitmap.
        using (Bitmap previewBitmap = new Bitmap(previewWidth, previewHeight))
        {
            using (Graphics graphics = Graphics.FromImage(previewBitmap))
            {
                graphics.Clear(Color.LightBlue);
                // Additional drawing (e.g., text) can be added here if desired.
            }

            // Save the bitmap as PNG.
            previewBitmap.Save(previewPath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document and insert the preview image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(previewPath);
        string docPath = Path.Combine(outputDir, "AudioDoc.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images from embedded audio tracks.
        //    (In this example the preview image is treated as the audio track image.)
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        int thumbnailCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image type.
                string imageExtension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedImagePath = Path.Combine(outputDir, $"extracted_{imageIndex}{imageExtension}");

                // Save the original image to a temporary file.
                shape.ImageData.Save(extractedImagePath);

                // -----------------------------------------------------------------
                // 4. Create a JPEG thumbnail (100x100) from the extracted image.
                // -----------------------------------------------------------------
                using (Image originalImage = Image.FromFile(extractedImagePath))
                {
                    int thumbSize = 100;
                    using (Bitmap thumbnailBitmap = new Bitmap(thumbSize, thumbSize))
                    {
                        using (Graphics graphics = Graphics.FromImage(thumbnailBitmap))
                        {
                            graphics.Clear(Color.White);
                            graphics.DrawImage(originalImage, new Rectangle(0, 0, thumbSize, thumbSize));
                        }

                        string thumbnailPath = Path.Combine(outputDir, $"AudioThumbnail_{imageIndex}.jpg");
                        thumbnailBitmap.Save(thumbnailPath, ImageFormat.Jpeg);
                        thumbnailCount++;
                    }
                }

                imageIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 5. Validate that at least one thumbnail was created.
        // -----------------------------------------------------------------
        if (thumbnailCount == 0)
            throw new InvalidOperationException("No thumbnails were created from the document.");

        Console.WriteLine($"Successfully created {thumbnailCount} thumbnail(s) in the folder: {outputDir}");
    }
}
