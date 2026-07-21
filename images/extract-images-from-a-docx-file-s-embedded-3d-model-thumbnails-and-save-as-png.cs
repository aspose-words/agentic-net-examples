using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define deterministic file names.
        const string thumbnailPath = "thumbnail.png";
        const string docPath = "sample.docx";
        const string outputFolder = "ExtractedImages";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // 1. Create a sample PNG image to act as a 3D model thumbnail.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with a solid color.
                g.Clear(Color.LightBlue);
                // Draw a simple ellipse to make the image recognizable.
                g.DrawEllipse(new Pen(Color.DarkBlue, 5), 20, 20, imgWidth - 40, imgHeight - 40);
            }
            // Save the bitmap as PNG.
            bitmap.Save(thumbnailPath);
        }

        // -------------------------------------------------
        // 2. Create a DOCX document and insert the thumbnail image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image; this creates a Shape with HasImage = true.
        builder.InsertImage(thumbnailPath);
        // Save the document.
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all images (thumbnails) as PNG files.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                // Force PNG extension regardless of original type (requirement).
                string outputFile = Path.Combine(outputFolder, $"extracted_{extractedCount}.png");

                // Save the image data to a temporary file using its original format.
                string tempFile = Path.Combine(outputFolder, $"temp_{extractedCount}{extension}");
                shape.ImageData.Save(tempFile);

                // If the original format is not PNG, convert it to PNG using Aspose.Drawing.
                if (!extension.Equals(".png", StringComparison.OrdinalIgnoreCase))
                {
                    using (var srcImage = Aspose.Drawing.Image.FromFile(tempFile))
                    {
                        srcImage.Save(outputFile, ImageFormat.Png);
                    }
                    File.Delete(tempFile);
                }
                else
                {
                    // Already PNG; just rename to the deterministic name.
                    File.Move(tempFile, outputFile);
                }

                extractedCount++;
            }
        }

        // -------------------------------------------------
        // 4. Validate that at least one image was extracted.
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // The program finishes automatically; no user interaction required.
    }
}
