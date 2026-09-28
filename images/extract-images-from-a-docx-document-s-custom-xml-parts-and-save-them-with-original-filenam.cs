using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image.
        string imagePath = "sample.png";
        CreateSampleImage(imagePath);

        // Step 2: Create a DOCX document and embed the image.
        string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, imagePath);

        // Step 3: Extract images from the document and save them with original filenames.
        ExtractImagesFromDocument(docPath);
    }

    private static void CreateSampleImage(string path)
    {
        int width = 200;
        int height = 100;

        // Create bitmap using Aspose.Drawing.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }

        // Save the bitmap to a deterministic file name.
        bitmap.Save(path);

        // Clean up drawing resources.
        graphics.Dispose();
        bitmap.Dispose();

        // Validate that the image file was created.
        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at '{path}'.");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        // Load the image bytes to obtain the original file name.
        string fileName = Path.GetFileName(imagePath);

        // Create a new empty document.
        Document doc = new Document();

        // Insert the image using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.InsertImage(imagePath);
        // Store the original file name in the shape's Title property for later extraction.
        shape.Title = fileName;

        // Save the document.
        doc.Save(docPath);

        // Validate that the document was saved.
        if (!File.Exists(docPath))
            throw new Exception($"Failed to create document at '{docPath}'.");
    }

    private static void ExtractImagesFromDocument(string docPath)
    {
        Document doc = new Document(docPath);
        int extractedCount = 0;
        int index = 1;

        // Iterate over all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Determine the output file name.
            string fileName = !string.IsNullOrEmpty(shape.Title)
                ? shape.Title
                : $"extracted-{index}.png";

            // Save the image data to a file.
            shape.ImageData.Save(fileName);
            extractedCount++;
            index++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        Console.WriteLine($"Successfully extracted {extractedCount} image(s) from the document.");
    }
}
