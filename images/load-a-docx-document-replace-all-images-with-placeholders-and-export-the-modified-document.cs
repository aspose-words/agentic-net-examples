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
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample image that will be inserted into the document
        string inputImagePath = Path.Combine(artifactsDir, "input.png");
        CreateSampleImage(inputImagePath, 200, 100, Color.LightBlue, "Sample");

        // Create a placeholder image that will replace existing images
        string placeholderImagePath = Path.Combine(artifactsDir, "placeholder.png");
        CreateSampleImage(placeholderImagePath, 200, 100, Color.LightGray, "Placeholder");

        // Build a sample DOCX containing a couple of images
        string originalDocPath = Path.Combine(artifactsDir, "original.docx");
        BuildSampleDocument(originalDocPath, inputImagePath);

        // Load the document
        Document doc = new Document(originalDocPath);

        // Replace every image with the placeholder image
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                // Ensure the shape really contains an image before accessing ImageData
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // Save the modified document
        string modifiedDocPath = Path.Combine(artifactsDir, "modified.docx");
        doc.Save(modifiedDocPath);

        // Validate that the output file was created
        if (!File.Exists(modifiedDocPath))
            throw new Exception("The modified document was not saved.");

        // Indicate success (no interactive prompts required)
        Console.WriteLine("Document processed successfully. Output: " + modifiedDocPath);
    }

    // Creates a deterministic bitmap with optional text and saves it to a file
    private static void CreateSampleImage(string filePath, int width, int height, Color backColor, string text)
    {
        // Explicitly use Aspose.Drawing types to avoid ambiguity
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
                // Draw simple text in the centre
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 16))
                {
                    SizeF textSize = g.MeasureString(text, font);
                    float x = (width - textSize.Width) / 2;
                    float y = (height - textSize.Height) / 2;
                    g.DrawString(text, font, Brushes.Black, x, y);
                }
            }
            // Save the bitmap as PNG
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Builds a sample document with two images inserted
    private static void BuildSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample document with images:");
        builder.InsertImage(imagePath);
        builder.Writeln("Some text between images.");
        builder.InsertImage(imagePath);
        builder.Writeln("End of document.");

        doc.Save(docPath);
    }
}
