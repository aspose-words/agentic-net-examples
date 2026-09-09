using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using System.Linq;

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        const string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample JPEG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        const string sampleImagePath = "sample.jpg";
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(Aspose.Drawing.Color.LightBlue);

                // Draw a simple ellipse.
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Orange))
                {
                    graphics.FillEllipse(brush, 20, 20, 160, 160);
                }
            }

            // Save the bitmap as a JPEG file.
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Create a DOCX document and insert the JPEG image several times.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.InsertParagraph();
        builder.InsertImage(sampleImagePath);

        // Save the document to the output folder.
        string docPath = Path.Combine(outputDir, "input.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and process all JPEG images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Apply grayscale filter to the image.
            shape.ImageData.GrayScale = true;

            // Determine the appropriate file extension for the image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outputImagePath = Path.Combine(outputDir, $"extracted_{imageIndex}{extension}");

            // Save the processed image to disk.
            shape.ImageData.Save(outputImagePath);
            imageIndex++;
        }

        // Validate that at least one image was extracted and processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found and processed.");

        // Optional cleanup of the temporary sample image.
        // File.Delete(sampleImagePath);
    }
}
