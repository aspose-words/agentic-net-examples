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
        // Create a deterministic sample image to be inserted into the document.
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath, 200, 150, Aspose.Drawing.Color.LightBlue, Aspose.Drawing.Color.DarkBlue);

        // Create a new document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string inputDocPath = "input.docx";
        doc.Save(inputDocPath);

        // Load the document that contains images.
        Document loadedDoc = new Document(inputDocPath);

        // Create a placeholder image that will replace all existing images.
        const string placeholderImagePath = "placeholder.png";
        CreateSampleImage(placeholderImagePath, 200, 150, Aspose.Drawing.Color.LightGray, Aspose.Drawing.Color.DarkGray);

        // Replace each image in the document with the placeholder image.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // Save the modified document.
        const string outputDocPath = "output.docx";
        loadedDoc.Save(outputDocPath);

        // Validate that the output file was created.
        if (!File.Exists(outputDocPath))
        {
            throw new Exception($"Failed to create the output document: {outputDocPath}");
        }

        Console.WriteLine("Image replacement completed successfully.");
    }

    // Helper method to create a simple deterministic PNG image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color background, Aspose.Drawing.Color border)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(background);

                // Draw a simple border rectangle.
                using (Pen pen = new Pen(border, 5))
                {
                    graphics.DrawRectangle(pen, 0, 0, width - 1, height - 1);
                }
            }

            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Ensure the file exists.
        if (!File.Exists(filePath))
        {
            throw new Exception($"Failed to create image file: {filePath}");
        }
    }
}
