using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Bitmap, Graphics, Color
using Aspose.Drawing.Imaging;      // ImageFormat

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string placeholderPath = Path.Combine(Directory.GetCurrentDirectory(), "placeholder.png");
        string sampleImagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        string inputDocPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputDocPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // 1. Create a placeholder image (100x100, light gray background).
        // -----------------------------------------------------------------
        CreateSampleImage(placeholderPath, 100, 100, Aspose.Drawing.Color.LightGray);

        // -----------------------------------------------------------------
        // 2. Create a sample image to be inserted into the document.
        // -----------------------------------------------------------------
        CreateSampleImage(sampleImagePath, 200, 150, Aspose.Drawing.Color.CornflowerBlue);

        // -----------------------------------------------------------------
        // 3. Build a sample DOCX containing a few images.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three images using the sample image file.
        builder.Writeln("First image:");
        builder.InsertImage(sampleImagePath);
        builder.Writeln();

        builder.Writeln("Second image:");
        builder.InsertImage(sampleImagePath);
        builder.Writeln();

        builder.Writeln("Third image:");
        builder.InsertImage(sampleImagePath);
        builder.Writeln();

        // Save the original document.
        doc.Save(inputDocPath);

        // -----------------------------------------------------------------
        // 4. Load the document from file.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);

        // -----------------------------------------------------------------
        // 5. Replace every image with the placeholder image.
        // -----------------------------------------------------------------
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                // Replace the image data with the placeholder.
                shape.ImageData.SetImage(placeholderPath);
            }
        }

        // -----------------------------------------------------------------
        // 6. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputDocPath);

        // -----------------------------------------------------------------
        // 7. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The output document was not created.");

        Console.WriteLine("Images replaced and document saved to: " + outputDocPath);
    }

    // Helper method to create a deterministic PNG image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        // Ensure any existing file is overwritten.
        if (File.Exists(filePath))
            File.Delete(filePath);

        // Create bitmap and fill with the specified background color.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(backgroundColor);

        // Save the bitmap to PNG format.
        bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);

        // Clean up resources.
        graphics.Dispose();
        bitmap.Dispose();
    }
}
