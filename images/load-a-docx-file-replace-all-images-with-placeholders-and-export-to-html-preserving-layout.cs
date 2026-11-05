using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json; // Included as per required packages

public class Program
{
    public static void Main()
    {
        // Define deterministic file names
        const string inputImagePath = "input.png";
        const string placeholderImagePath = "placeholder.png";
        const string inputDocPath = "input.docx";
        const string outputHtmlPath = "output.html";

        // Ensure a clean environment
        DeleteIfExists(inputImagePath);
        DeleteIfExists(placeholderImagePath);
        DeleteIfExists(inputDocPath);
        DeleteIfExists(outputHtmlPath);

        // 1. Create a sample image to be inserted into the document
        CreateSampleImage(inputImagePath, 300, 200, Aspose.Drawing.Color.LightBlue);

        // 2. Create a placeholder image that will replace existing images
        CreateSampleImage(placeholderImagePath, 300, 200, Aspose.Drawing.Color.LightGray);

        // 3. Create a DOCX file containing the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with an image:");
        builder.InsertImage(inputImagePath);
        doc.Save(inputDocPath);

        // 4. Load the DOCX file
        Document loadedDoc = new Document(inputDocPath);

        // 5. Replace all images with the placeholder image
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // 6. Export the modified document to HTML while preserving layout
        loadedDoc.Save(outputHtmlPath, SaveFormat.Html);

        // 7. Validate that the HTML output was created
        if (!File.Exists(outputHtmlPath))
            throw new Exception("HTML output file was not created.");

        // Optional: Indicate success (no interactive prompts)
        Console.WriteLine("Processing completed successfully.");
    }

    private static void CreateSampleImage(string path, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
                // Draw a simple border
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 5))
                {
                    graphics.DrawRectangle(pen, 0, 0, width - 1, height - 1);
                }
            }
            // Save as PNG
            bitmap.Save(path, ImageFormat.Png);
        }

        // Validate that the image file exists
        if (!File.Exists(path))
            throw new Exception($"Failed to create image file: {path}");
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }
}
