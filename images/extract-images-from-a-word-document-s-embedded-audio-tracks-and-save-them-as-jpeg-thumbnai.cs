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
        // Step 1: Create a deterministic sample image (input.png)
        const string sampleImagePath = "input.png";
        using (var bitmap = new Bitmap(200, 200))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.DrawRectangle(new Pen(Color.Blue, 5), 20, 20, 160, 160);
            }
            bitmap.Save(sampleImagePath);
        }

        // Step 2: Create a Word document and embed the sample image (simulating an audio thumbnail)
        const string docPath = "sample.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Document with an embedded audio thumbnail (image).");
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Step 3: Load the document and extract images from shapes (audio tracks simulation)
        var loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        const string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Save original image to a memory stream
            using (var originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into Aspose.Drawing.Bitmap
                using (var originalBitmap = new Bitmap(originalStream))
                {
                    // Create thumbnail bitmap (e.g., 100x100)
                    const int thumbWidth = 100;
                    const int thumbHeight = 100;
                    using (var thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (var g = Graphics.FromImage(thumbBitmap))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(originalBitmap, new Rectangle(0, 0, thumbWidth, thumbHeight));
                        }

                        // Save thumbnail as JPEG
                        imageIndex++;
                        string thumbPath = Path.Combine(outputFolder, $"thumbnail-{imageIndex}.jpg");
                        thumbBitmap.Save(thumbPath, ImageFormat.Jpeg);
                    }
                }
            }
        }

        // Validation: ensure at least one thumbnail was created
        var createdFiles = Directory.GetFiles(outputFolder, "thumbnail-*.jpg");
        if (createdFiles.Length == 0)
            throw new InvalidOperationException("No thumbnails were extracted.");
    }
}
