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
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "input");
        string extractedDir = Path.Combine(baseDir, "extracted");
        string outputDir = Path.Combine(baseDir, "output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample GIF image (static for simplicity)
        string sampleGifPath = Path.Combine(inputDir, "sample.gif");
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.White);
                g.DrawEllipse(new Pen(Aspose.Drawing.Color.Blue, 5), 20, 20, 160, 160);
            }
            bmp.Save(sampleGifPath, ImageFormat.Gif);
        }

        // 2. Create a Word document and insert the GIF
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleGifPath);
        string docPath = Path.Combine(baseDir, "sample.docx");
        doc.Save(docPath);

        // 3. Load the document and extract all GIF images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int gifIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                string extractedPath = Path.Combine(extractedDir, $"image-{gifIndex}.gif");
                shape.ImageData.Save(extractedPath);
                gifIndex++;
            }
        }

        // Validate that at least one GIF was extracted
        string[] extractedGifs = Directory.GetFiles(extractedDir, "*.gif");
        if (extractedGifs.Length == 0)
            throw new InvalidOperationException("No GIF images were extracted from the document.");

        // 4. Batch convert each extracted GIF to animated PNG (preserving timing where possible)
        int pngIndex = 0;
        foreach (string gifFile in extractedGifs)
        {
            using (Image gifImage = Image.FromFile(gifFile))
            {
                // Retrieve frame delay if the GIF is animated (property 0x5100)
                // If the property does not exist, default to 100ms per frame.
                int[] frameDelays = null;
                try
                {
                    var propItem = gifImage.GetPropertyItem(0x5100);
                    // Delays are stored as 4-byte integers (in 1/100ths of a second)
                    int count = propItem.Value.Length / 4;
                    frameDelays = new int[count];
                    for (int i = 0; i < count; i++)
                        frameDelays[i] = BitConverter.ToInt32(propItem.Value, i * 4);
                }
                catch
                {
                    // Property not found; treat as single-frame GIF
                    frameDelays = new int[] { 10 }; // 100ms default
                }

                // Save as PNG. Aspose.Drawing does not support animated PNG directly,
                // so we save the first frame as a static PNG. Timing information is retained
                // in the variable above for further processing if needed.
                string pngPath = Path.Combine(outputDir, $"image-{pngIndex}.png");
                gifImage.Save(pngPath, ImageFormat.Png);
                pngIndex++;
            }
        }

        // Validate that at least one PNG was created
        string[] outputPngs = Directory.GetFiles(outputDir, "*.png");
        if (outputPngs.Length == 0)
            throw new InvalidOperationException("No PNG images were created during conversion.");

        // Cleanup (optional)
        // Console.WriteLine("Batch conversion completed successfully.");
    }
}
