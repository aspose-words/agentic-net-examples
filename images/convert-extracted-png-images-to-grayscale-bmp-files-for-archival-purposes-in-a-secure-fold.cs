using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class ConvertPngToGrayscaleBmp
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string secureDir = Path.Combine(artifactsDir, "SecureFolder");
        Directory.CreateDirectory(secureDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic PNG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        int width = 200;
        int height = 100;
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 3))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }
        string pngPath = "input.png";
        bitmap.Save(pngPath);
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Insert the PNG image into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath); // insertion strict rule

        // -----------------------------------------------------------------
        // 3. Extract images, convert each to grayscale BMP, and save.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int savedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue; // extraction strict rule

            // Force grayscale rendering.
            shape.ImageData.GrayScale = true;

            // Save as BMP in the secure folder.
            string bmpFileName = Path.Combine(secureDir, $"Image{savedCount}.bmp");
            shape.ImageData.Save(bmpFileName);
            savedCount++;
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one image was saved.
        // -----------------------------------------------------------------
        if (savedCount == 0)
            throw new InvalidOperationException("No images were extracted and saved.");

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine($"Successfully saved {savedCount} grayscale BMP image(s) to '{secureDir}'.");
    }
}
