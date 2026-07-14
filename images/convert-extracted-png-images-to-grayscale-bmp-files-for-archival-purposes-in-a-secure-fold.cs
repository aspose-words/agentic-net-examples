using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Provides Bitmap, Graphics, Color
using Aspose.Drawing.Imaging;      // Provides ImageFormat

public class Program
{
    public static void Main()
    {
        // Define deterministic file and folder names
        string workingDir = Directory.GetCurrentDirectory();
        string inputImagePath = Path.Combine(workingDir, "input.png");
        string secureFolder = Path.Combine(workingDir, "SecureFolder");
        Directory.CreateDirectory(secureFolder);

        // -------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white and draw a simple rectangle
                g.Clear(Color.White);
                g.FillRectangle(new SolidBrush(Color.Blue), 20, 20, 160, 60);
            }

            // Save the bitmap as PNG – this will be the source image for the document
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the PNG image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);

        // -------------------------------------------------
        // 3. Extract all images, convert each to grayscale BMP,
        //    and save them into the secure folder
        // -------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int savedCount = 0;
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Apply grayscale rendering
            shape.ImageData.GrayScale = true;

            // Build deterministic output file name
            string outputPath = Path.Combine(secureFolder, $"extracted_{imageIndex}.bmp");

            // Save the image as BMP
            shape.ImageData.Save(outputPath);
            savedCount++;
            imageIndex++;
        }

        // -------------------------------------------------
        // 4. Validation – ensure at least one image was saved
        // -------------------------------------------------
        if (savedCount == 0)
            throw new InvalidOperationException("No images were extracted and saved.");

        Console.WriteLine($"Successfully extracted and saved {savedCount} grayscale BMP image(s) to '{secureFolder}'.");
    }
}
