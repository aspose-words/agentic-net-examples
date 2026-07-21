using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchTiffToJpegConverter
{
    public static void Main()
    {
        // Prepare input and output directories
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create deterministic sample TIFF images
        CreateSampleTiff(Path.Combine(inputDir, "sample1.tif"), 200, 200, Color.LightBlue);
        CreateSampleTiff(Path.Combine(inputDir, "sample2.tif"), 150, 150, Color.LightCoral);

        // Build a Word document that contains the sample TIFF images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with TIFF images:");
        builder.InsertImage(Path.Combine(inputDir, "sample1.tif"));
        builder.InsertParagraph();
        builder.InsertImage(Path.Combine(inputDir, "sample2.tif"));
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDoc.docx");
        doc.Save(docPath);

        // Load the document and locate all shapes that contain images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        var imageShapes = shapeNodes
            .OfType<Shape>()
            .Where(s => s.HasImage) // Process any image shape (TIFF images appear as Unknown type)
            .ToList();

        if (!imageShapes.Any())
            throw new InvalidOperationException("No images were found for conversion.");

        int index = 0;
        foreach (Shape shape in imageShapes)
        {
            // Save the extracted image to a temporary file
            string tempPath = Path.Combine(inputDir, $"extracted_{index}.tif");
            using (FileStream tempStream = File.Create(tempPath))
                shape.ImageData.Save(tempStream);

            // Create a new document that contains only this image
            Document singleImageDoc = new Document();
            DocumentBuilder imgBuilder = new DocumentBuilder(singleImageDoc);
            imgBuilder.InsertImage(tempPath);

            // Configure JPEG save options with 90% quality
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 90
            };

            // Save the image as JPEG
            string outputJpegPath = Path.Combine(outputDir, $"converted_{index}.jpg");
            singleImageDoc.Save(outputJpegPath, jpegOptions);

            // Verify that the JPEG file was created
            if (!File.Exists(outputJpegPath))
                throw new InvalidOperationException($"Failed to create JPEG file: {outputJpegPath}");

            index++;
        }

        // Optional cleanup of temporary TIFF files
        foreach (string tempFile in Directory.GetFiles(inputDir, "extracted_*.tif"))
            File.Delete(tempFile);
    }

    // Helper method to create a deterministic TIFF image using Aspose.Drawing
    private static void CreateSampleTiff(string filePath, int width, int height, Color backgroundColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backgroundColor);
            using (Pen pen = new Pen(Color.Black, 3))
            {
                graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath, ImageFormat.Tiff);
        }
    }
}
