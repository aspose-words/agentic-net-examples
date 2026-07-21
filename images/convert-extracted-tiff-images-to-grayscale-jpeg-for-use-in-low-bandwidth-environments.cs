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
        // Directories for input and output files
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string inputImagePath = Path.Combine(artifactsDir, "sample.tiff");
        string docPath = Path.Combine(artifactsDir, "document.docx");

        // -------------------------------------------------
        // 1. Create a deterministic sample TIFF image
        // -------------------------------------------------
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white
                g.Clear(Color.White);
                // Draw a simple red rectangle
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 20, 20, width - 40, height - 40);
                }
            }
            // Save as TIFF
            bitmap.Save(inputImagePath, ImageFormat.Tiff);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the TIFF image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract images,
        //    convert each to grayscale JPEG, and save.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Set the shape's image to display in grayscale (affects rendering)
            shape.ImageData.GrayScale = true;

            // Prepare output file name
            string outputFilePath = Path.Combine(artifactsDir,
                $"grayscale_image_{imageIndex}.jpg");

            // Configure JPEG save options with grayscale color mode
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 80,
                ImageColorMode = ImageColorMode.Grayscale
            };

            // Save the shape via a temporary document to apply the rendering options
            Document tempDoc = new Document();
            DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
            Shape clonedShape = (Shape)tempDoc.ImportNode(shape, true);
            tempBuilder.InsertNode(clonedShape);
            tempDoc.Save(outputFilePath, jpegOptions);

            // Validate that the file was created
            if (!File.Exists(outputFilePath))
                throw new InvalidOperationException($"Failed to create {outputFilePath}");

            imageIndex++;
        }

        // If no images were processed, raise an error
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found to convert.");
    }
}
