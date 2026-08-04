using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample TIFF image using Aspose.Drawing
        string tiffPath = Path.Combine(artifactsDir, "sample.tif");
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red))
                    {
                        g.DrawString("TIFF", font, brush, new Aspose.Drawing.PointF(50, 80));
                    }
                }
            }
            bitmap.Save(tiffPath, Aspose.Drawing.Imaging.ImageFormat.Tiff);
        }

        // 2. Insert the TIFF image into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(tiffPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithTiff.docx");
        doc.Save(docPath);

        // 3. Load the document and extract images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue; // Skip shapes without images

            // Save the image to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // 4. Create a temporary document containing only this image
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.InsertImage(imageStream);

                // 5. Save the temporary document as a grayscale JPEG
                ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
                {
                    ImageColorMode = ImageColorMode.Grayscale,
                    JpegQuality = 80
                };

                string jpegPath = Path.Combine(artifactsDir, $"Image_{imageIndex}_grayscale.jpg");
                tempDoc.Save(jpegPath, jpegOptions);

                // Validate that the JPEG file was created
                if (!File.Exists(jpegPath))
                    throw new InvalidOperationException($"Failed to create JPEG file: {jpegPath}");

                imageIndex++;
            }
        }

        // Ensure at least one image was processed
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found to convert.");
    }
}
