using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image.
        string imagePath = "sample.png";
        CreateSampleImage(imagePath);

        // Create a Word document that contains an equation and the sample image.
        string docPath = "sample.docx";
        CreateDocumentWithEquationAndImage(docPath, imagePath);

        // Extract all images that are stored inside the document (including those inside equations).
        ExtractEquationImages(docPath);
    }

    private static void CreateSampleImage(string path)
    {
        int width = 100;
        int height = 100;

        // Use Aspose.Drawing to create a simple bitmap.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);

        // Fill background and draw a rectangle.
        graphics.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }

        // Save the bitmap to a file.
        bitmap.Save(path);

        // Clean up resources.
        graphics.Dispose();
        bitmap.Dispose();

        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at '{path}'.");
    }

    private static void CreateDocumentWithEquationAndImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple equation field.
        builder.Writeln("Equation with embedded image:");
        builder.InsertField("EQ \\o(\\a\\b,\\c\\d)");

        // Insert the image after the equation.
        builder.Writeln();
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new Exception($"Failed to create document at '{docPath}'.");
    }

    private static void ExtractEquationImages(string docPath)
    {
        Document doc = new Document(docPath);

        // Retrieve all Shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Ensure the shape actually contains an image.
            if (shape.HasImage)
            {
                string outFile = $"equation-image-{extractedCount}.png";
                shape.ImageData.Save(outFile);
                if (!File.Exists(outFile))
                    throw new Exception($"Failed to save extracted image to '{outFile}'.");
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        Console.WriteLine($"{extractedCount} image(s) extracted from the document.");
    }
}
