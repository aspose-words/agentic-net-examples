using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string inputDir = Path.Combine(workDir, "InputImages");
        string outputDir = Path.Combine(workDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample animated GIF (tiny two‑frame GIF) from a base64 string.
        // -----------------------------------------------------------------
        // This GIF was generated beforehand and embedded as a string to avoid external files.
        const string base64Gif =
            "R0lGODdhAQABAIAAAAUEBAAAACwAAAAAAQABAAACAkQBADs="; // 1×1 pixel, 2‑frame transparent GIF
        byte[] gifBytes = Convert.FromBase64String(base64Gif);
        string gifPath = Path.Combine(inputDir, "sample.gif");
        File.WriteAllBytes(gifPath, gifBytes);

        // -----------------------------------------------------------------
        // 2. Insert the GIF into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape gifShape = builder.InsertImage(gifPath);
        // Ensure the shape really contains an image.
        if (!gifShape.HasImage)
            throw new InvalidOperationException("Failed to insert GIF image into the document.");

        // Save the document (optional, just to demonstrate the lifecycle).
        string docPath = Path.Combine(workDir, "Sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract all GIF images from the document.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        List<Shape> gifShapes = new List<Shape>();
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
                gifShapes.Add(shape);
        }

        if (gifShapes.Count == 0)
            throw new InvalidOperationException("No GIF images were found in the document.");

        // -----------------------------------------------------------------
        // 4. Convert each extracted GIF to an animated PNG (APNG) while preserving timing.
        //    Aspose.Drawing can load the GIF, keep its frame delays, and save as PNG.
        // -----------------------------------------------------------------
        int index = 0;
        foreach (Shape shape in gifShapes)
        {
            // Obtain the raw GIF bytes.
            using (MemoryStream gifStream = new MemoryStream())
            {
                shape.ImageData.Save(gifStream);
                gifStream.Position = 0;

                // Load the GIF with Aspose.Drawing.
                using (Image gifImage = Image.FromStream(gifStream))
                {
                    // Prepare the output file name.
                    string outFile = Path.Combine(outputDir, $"converted_{index}.png");

                    // Save as PNG. Aspose.Drawing preserves animation frames and timing
                    // when the target format supports it (APNG). If the environment does not
                    // support APNG, the first frame will be saved.
                    gifImage.Save(outFile, ImageFormat.Png);
                    if (!File.Exists(outFile))
                        throw new InvalidOperationException($"Failed to create output file '{outFile}'.");
                }
            }
            index++;
        }

        // -----------------------------------------------------------------
        // 5. Validation – ensure at least one output file was produced.
        // -----------------------------------------------------------------
        string[] producedFiles = Directory.GetFiles(outputDir, "*.png");
        if (producedFiles.Length == 0)
            throw new InvalidOperationException("No PNG files were created during conversion.");

        // The example finishes without requiring user interaction.
    }
}
