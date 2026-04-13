using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image.
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath);

        // Create a PDF document that contains the sample image.
        const string pdfPath = "sample.pdf";
        CreatePdfWithImage(sampleImagePath, pdfPath);

        // Load the PDF, extract embedded images and convert them to JPEG with 85% quality.
        ExtractImagesToJpeg(pdfPath);
    }

    private static void CreateSampleImage(string filePath)
    {
        int width = 200;
        int height = 100;

        // Create a bitmap and draw simple content.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        using (Pen pen = new Pen(Color.Blue, 3))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }

        // Save the bitmap to a file.
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();

        if (!File.Exists(filePath))
            throw new Exception("Failed to create the sample image.");
    }

    private static void CreatePdfWithImage(string imagePath, string pdfPath)
    {
        // Build a simple document and insert the image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);

        // Save the document as PDF.
        doc.Save(pdfPath);
        if (!File.Exists(pdfPath))
            throw new Exception("Failed to create the PDF file.");
    }

    private static void ExtractImagesToJpeg(string pdfPath)
    {
        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Collect all shape nodes.
        NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the shape's image data to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reuse.

                // Insert the extracted image into a temporary document.
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.InsertImage(imageStream);

                // Configure JPEG output with 85% quality.
                ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
                {
                    JpegQuality = 85
                };

                // Save the image as a JPEG file.
                string outputFile = $"extracted_{imageIndex}.jpg";
                tempDoc.Save(outputFile, jpegOptions);

                if (!File.Exists(outputFile))
                    throw new Exception($"Failed to save extracted image '{outputFile}'.");

                imageIndex++;
            }
        }

        if (imageIndex == 0)
            throw new Exception("No images were extracted from the PDF.");
    }
}
