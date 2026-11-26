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
        // Prepare folders
        string inputFolder = "InputDocs";
        string outputFolder = "ExtractedImages";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = "input.png";
        CreateSampleImage(sampleImagePath);

        // Create sample DOCX files containing the image
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleImagePath);

        // Batch process each DOCX file in the input folder
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (!shape.HasImage)
                    continue;

                // Extract image to a memory stream
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0; // Reset before reading

                    // Load the image with Aspose.Drawing and save as BMP
                    using (Image drawingImage = Image.FromStream(imageStream))
                    {
                        string outputFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}.bmp";
                        string outputPath = Path.Combine(outputFolder, outputFileName);
                        drawingImage.Save(outputPath, ImageFormat.Bmp);

                        // Validate that the file was created
                        if (!File.Exists(outputPath))
                            throw new InvalidOperationException($"Failed to save extracted image to '{outputPath}'.");
                    }
                }

                imageIndex++;
            }

            // Ensure at least one image was extracted from the document
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were found in document '{docPath}'.");
        }

        // Final validation: at least one BMP file should exist in the output folder
        if (Directory.GetFiles(outputFolder, "*.bmp").Length == 0)
            throw new InvalidOperationException("No BMP images were extracted from the batch processing.");

        // Example completed without interactive prompts
    }

    private static void CreateSampleImage(string path)
    {
        const int width = 200;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle for visual distinction
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path);
        }

        // Validate that the image file exists
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample image at '{path}'.");
    }

    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // Validate that the document file exists
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create sample document at '{docPath}'.");
    }
}
