using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

namespace ExtractVideoThumbnails
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a folder for all generated files.
            string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            Directory.CreateDirectory(dataDir);

            // Create a sample thumbnail image using Aspose.Drawing.
            string thumbnailPath = Path.Combine(dataDir, "thumb.png");
            const int width = 200;
            const int height = 150;
            using (Bitmap bitmap = new Bitmap(width, height))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple red rectangle to make the thumbnail recognizable.
                graphics.FillRectangle(new SolidBrush(Color.Red), 20, 20, width - 40, height - 40);
                bitmap.Save(thumbnailPath);
            }

            // Create a new Word document and insert the thumbnail as if it were a video preview.
            string docPath = Path.Combine(dataDir, "VideoDoc.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Insert the image; in a real scenario this would be the video thumbnail.
            Shape videoShape = builder.InsertImage(thumbnailPath);
            // Optionally mark the shape to indicate it represents a video.
            videoShape.Title = "VideoThumbnail";

            // Save the document.
            doc.Save(docPath);

            // Load the document and extract all images (thumbnails) from shapes.
            Document loadedDoc = new Document(docPath);
            NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine the appropriate file extension based on the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outputPath = Path.Combine(dataDir, $"ExtractedThumbnail{imageIndex}{extension}");
                    shape.ImageData.Save(outputPath);
                    imageIndex++;
                }
            }

            // Validate that at least one image was extracted.
            if (imageIndex == 0)
                throw new InvalidOperationException("No thumbnail images were extracted from the document.");

            // The example finishes without requiring user interaction.
        }
    }
}
