using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string outputDir = Path.Combine(artifactsDir, "Output");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample TIFF image
        // -----------------------------------------------------------------
        string tiffPath = Path.Combine(artifactsDir, "sample.tif");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple blue rectangle
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    g.FillRectangle(brush, 20, 20, 160, 160);
                }
            }
            // Save as TIFF
            bitmap.Save(tiffPath, ImageFormat.Tiff);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the TIFF image
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with a TIFF image:");
        builder.InsertImage(tiffPath);
        string docPath = Path.Combine(artifactsDir, "DocWithTiff.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue; // Ensure the shape actually contains an image

            // -----------------------------------------------------------------
            // 4. Convert each extracted image to JPEG with 90% quality
            // -----------------------------------------------------------------
            using (MemoryStream imageStream = new MemoryStream())
            {
                // Save the original image data to a stream
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading

                // Create a temporary document that contains only this image
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.Writeln("Converted image:");
                tempBuilder.InsertImage(imageStream);

                // Configure JPEG save options with 90% quality
                ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
                {
                    JpegQuality = 90
                };

                // Save the temporary document as a JPEG image
                string outputFile = Path.Combine(outputDir, $"converted_{imageIndex}.jpg");
                tempDoc.Save(outputFile, jpegOptions);
                imageIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 5. Validation – ensure at least one JPEG was created
        // -----------------------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found and converted.");
    }
}
