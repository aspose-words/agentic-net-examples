using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample waveform image (PNG) locally.
        const string waveformImagePath = "waveform.png";
        CreateSampleWaveformImage(waveformImagePath);

        // Step 2: Create a DOCX document and insert the waveform image.
        const string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, waveformImagePath);

        // Step 3: Load the document and extract all embedded images, saving them as PNG files.
        ExtractImagesFromDocument(docPath);
    }

    private static void CreateSampleWaveformImage(string path)
    {
        const int width = 200;
        const int height = 100;

        // Create a bitmap and draw a simple placeholder for a waveform.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);

                // Draw a simple rectangle to represent the waveform area.
                using (Pen pen = new Pen(Color.Black, 2))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(path);
        }

        // Validate that the image file was created.
        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at '{path}'.");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        // Create a new empty document.
        Document doc = new Document();

        // Insert the image into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);

        // Validate that the document file was created.
        if (!File.Exists(docPath))
            throw new Exception($"Failed to create document at '{docPath}'.");
    }

    private static void ExtractImagesFromDocument(string docPath)
    {
        // Load the document that contains the embedded images.
        Document doc = new Document(docPath);

        // Collect all Shape nodes (which may contain images).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Save each extracted image as a PNG file with a deterministic name.
                string outputPath = $"extracted-{extractedCount}.png";
                shape.ImageData.Save(outputPath);

                // Validate that the image file was saved.
                if (!File.Exists(outputPath))
                    throw new Exception($"Failed to save extracted image to '{outputPath}'.");

                extractedCount++;
            }
        }

        // Ensure that at least one image was extracted.
        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");
    }
}
