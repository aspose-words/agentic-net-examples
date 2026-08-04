using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Define deterministic file and folder names.
        // -------------------------------------------------
        string artifactsDir = "Artifacts";
        string gifPath = "sample.gif";
        string pngPath = "sample.png";
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");

        // Ensure the output folder exists.
        Directory.CreateDirectory(artifactsDir);

        // -------------------------------------------------
        // 2. Create a sample GIF image.
        // -------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            g.FillRectangle(Brushes.Blue, 10, 10, 80, 80);
            bitmap.Save(gifPath, ImageFormat.Gif);
        }

        // -------------------------------------------------
        // 3. Create the equivalent PNG image.
        // -------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            g.FillRectangle(Brushes.Blue, 10, 10, 80, 80);
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 4. Build a Word document that contains the GIF image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with a GIF image:");
        builder.InsertImage(gifPath);
        doc.Save(inputDocPath);

        // -------------------------------------------------
        // 5. Define a custom mapping from GIF to PNG.
        // -------------------------------------------------
        var gifToPngMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { gifPath, pngPath }
        };

        // -------------------------------------------------
        // 6. Load the document and replace GIF images.
        // -------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        int replacedCount = 0;

        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // In this example we know the source file name.
                string sourceKey = gifPath;
                if (gifToPngMap.TryGetValue(sourceKey, out string replacementPath) && File.Exists(replacementPath))
                {
                    shape.ImageData.SetImage(replacementPath);
                    replacedCount++;
                }
            }
        }

        // -------------------------------------------------
        // 7. Save the modified document.
        // -------------------------------------------------
        loadedDoc.Save(outputDocPath);

        // -------------------------------------------------
        // 8. Validation.
        // -------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The output document was not created.");

        if (replacedCount == 0)
            throw new InvalidOperationException("No GIF images were replaced.");

        // Example completed without interactive input.
    }
}
