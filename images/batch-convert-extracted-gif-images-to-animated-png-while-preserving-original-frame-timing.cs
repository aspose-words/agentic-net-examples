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
        // Directories for input and output files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string inputDir = Path.Combine(workDir, "Input");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample GIF image (single‑frame for simplicity).
        // -----------------------------------------------------------------
        string gifPath = Path.Combine(inputDir, "sample.gif");
        using (Bitmap bmp = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.LightBlue);
            g.DrawEllipse(new Pen(Aspose.Drawing.Color.DarkRed, 5), 20, 20, 160, 160);
            // Save as GIF.
            bmp.Save(gifPath, ImageFormat.Gif);
        }

        // -----------------------------------------------------------------
        // 2. Insert the GIF into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document containing a GIF image:");
        builder.InsertImage(gifPath);
        string docPath = Path.Combine(workDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all GIF images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int gifIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Gif)
                continue;

            // -----------------------------------------------------------------
            // 4. Save the GIF image to a memory stream.
            // -----------------------------------------------------------------
            using (MemoryStream gifStream = new MemoryStream())
            {
                shape.ImageData.Save(gifStream);
                gifStream.Position = 0; // Reset before reading.

                // -----------------------------------------------------------------
                // 5. Load the GIF with Aspose.Drawing and save as PNG.
                //    (For a real animated GIF you would need to copy frame timing
                //     to an animated PNG; this example saves the first frame.)
                // -----------------------------------------------------------------
                using (Image gifImage = Image.FromStream(gifStream))
                {
                    string pngFileName = Path.Combine(outputDir, $"ExtractedGif_{gifIndex}.png");
                    gifImage.Save(pngFileName, ImageFormat.Png);

                    // Validate that the PNG file was created.
                    if (!File.Exists(pngFileName))
                        throw new InvalidOperationException($"Failed to create PNG file: {pngFileName}");
                }
            }

            gifIndex++;
        }

        // -----------------------------------------------------------------
        // 6. Final validation – at least one PNG should exist.
        // -----------------------------------------------------------------
        string[] pngFiles = Directory.GetFiles(outputDir, "*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG files were generated from GIF extraction.");

        // The program finishes without user interaction.
    }
}
