using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create deterministic sample image to insert into the document.
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath, 200, 100, "Sample");

        // Create a placeholder image that will replace original images.
        const string placeholderImagePath = "placeholder.png";
        CreateSampleImage(placeholderImagePath, 200, 100, "Placeholder");

        // Build a DOCX file containing the sample image.
        const string inputDocPath = "input.docx";
        CreateDocumentWithImage(inputDocPath, sampleImagePath);

        // Load the DOCX file.
        Document doc = new Document(inputDocPath);

        // Replace each image with the placeholder image while preserving layout.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // Export the modified document to HTML.
        const string outputHtmlPath = "output.html";
        doc.Save(outputHtmlPath, SaveFormat.Html);

        // Validate that the HTML file was created.
        if (!File.Exists(outputHtmlPath))
            throw new Exception("HTML output was not created.");
    }

    private static void CreateSampleImage(string path, int width, int height, string text)
    {
        // Ensure any existing file is removed.
        if (File.Exists(path))
            File.Delete(path);

        // Create bitmap.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            // Create graphics object.
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(Aspose.Drawing.Color.LightGray);

                // Draw text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                {
                    graphics.DrawString(text, font, brush, new Aspose.Drawing.PointF(10, height / 2 - 10));
                }
            }

            // Save bitmap to file.
            bitmap.Save(path, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // Validate that the image file exists.
        if (!File.Exists(path))
            throw new Exception($"Failed to create image file: {path}");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        // Ensure any existing file is removed.
        if (File.Exists(docPath))
            File.Delete(docPath);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document with an image:");
        builder.InsertImage(imagePath);

        doc.Save(docPath, SaveFormat.Docx);

        // Validate that the document file exists.
        if (!File.Exists(docPath))
            throw new Exception($"Failed to create document file: {docPath}");
    }
}
