using System;
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
        // Prepare folders.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string docsDir = Path.Combine(baseDir, "Docs");
        string outputDir = Path.Combine(baseDir, "SepiaImages");
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic PNG image to be used in the sample documents.
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSamplePng(sampleImagePath);

        // Create a few sample Word documents that contain the PNG image.
        const int documentCount = 2;
        for (int i = 1; i <= documentCount; i++)
        {
            string docPath = Path.Combine(docsDir, $"Document{i}.docx");
            CreateSampleDocument(docPath, sampleImagePath, i);
        }

        // Process each document: extract PNG images, apply sepia, and save.
        int totalProcessed = 0;
        int docNumber = 0;
        foreach (string docFile in Directory.GetFiles(docsDir, "*.docx"))
        {
            docNumber++;
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
                {
                    // Extract the image to a memory stream.
                    using (MemoryStream imageStream = new MemoryStream())
                    {
                        shape.ImageData.Save(imageStream);
                        imageStream.Position = 0; // Reset before reading.

                        // Load the image into a bitmap (original may be indexed).
                        using (Bitmap original = new Bitmap(imageStream))
                        {
                            // Create a non‑indexed bitmap to allow pixel manipulation.
                            using (Bitmap bitmap = new Bitmap(original.Width, original.Height, PixelFormat.Format24bppRgb))
                            {
                                using (Graphics g = Graphics.FromImage(bitmap))
                                {
                                    g.DrawImage(original, 0, 0, original.Width, original.Height);
                                }

                                // Apply sepia tone.
                                ApplySepia(bitmap);

                                // Save the modified image.
                                string outFile = Path.Combine(
                                    outputDir,
                                    $"sepia_doc{docNumber}_img{imageIndex}.png");
                                bitmap.Save(outFile, ImageFormat.Png);
                            }
                        }
                    }

                    imageIndex++;
                    totalProcessed++;
                }
            }
        }

        // Validation: ensure at least one PNG image was processed.
        if (totalProcessed == 0)
            throw new InvalidOperationException("No PNG images were found in the provided documents.");

        // Additional validation: ensure at least one output file exists.
        if (!Directory.GetFiles(outputDir, "*.png").Any())
            throw new InvalidOperationException("Sepia processing completed but no output images were created.");
    }

    // Creates a simple 100x100 PNG image with a solid background.
    private static void CreateSamplePng(string filePath)
    {
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a Word document that contains a paragraph and the provided PNG image.
    private static void CreateSampleDocument(string docPath, string imagePath, int docIndex)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Sample Document {docIndex}");
        builder.InsertImage(imagePath);
        doc.Save(docPath, SaveFormat.Docx);
    }

    // Applies a sepia tone effect to the supplied bitmap.
    private static void ApplySepia(Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                Aspose.Drawing.Color original = bitmap.GetPixel(x, y);
                int r = original.R;
                int g = original.G;
                int b = original.B;

                int tr = (int)(0.393 * r + 0.769 * g + 0.189 * b);
                int tg = (int)(0.349 * r + 0.686 * g + 0.168 * b);
                int tb = (int)(0.272 * r + 0.534 * g + 0.131 * b);

                tr = Math.Min(255, tr);
                tg = Math.Min(255, tg);
                tb = Math.Min(255, tb);

                Aspose.Drawing.Color sepia = Aspose.Drawing.Color.FromArgb(tr, tg, tb);
                bitmap.SetPixel(x, y, sepia);
            }
        }
    }
}
