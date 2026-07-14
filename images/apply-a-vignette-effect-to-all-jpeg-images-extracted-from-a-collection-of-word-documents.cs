using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string outputImagesDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(outputImagesDir);

        // Create a sample JPEG image to be inserted into documents
        string sampleJpegPath = Path.Combine(baseDir, "sample.jpg");
        CreateSampleJpeg(sampleJpegPath, 300, 200);

        // Create a few sample Word documents each containing the JPEG image
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputDocsDir, $"Document{i}.docx");
            CreateDocumentWithJpeg(docPath, sampleJpegPath);
        }

        // Process each document: extract JPEG images, apply vignette, save result
        int outputCount = 0;
        foreach (string docFile in Directory.GetFiles(inputDocsDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg);

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                // Save the original image to a memory stream
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0;

                    // Load the image into Aspose.Drawing.Bitmap
                    using (Bitmap bitmap = new Bitmap(imgStream))
                    {
                        // Apply vignette effect
                        ApplyVignette(bitmap);

                        // Save the processed image
                        string outFile = Path.Combine(outputImagesDir,
                            $"doc{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}.jpg");
                        bitmap.Save(outFile, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        outputCount++;
                    }
                }
                imageIndex++;
            }
        }

        // Validation: ensure at least one image was produced
        if (outputCount == 0)
            throw new InvalidOperationException("No JPEG images were extracted and processed.");

        // Example completed without interactive prompts
    }

    // Creates a deterministic JPEG image using Aspose.Drawing
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 100, 150, 200)))
            {
                g.FillRectangle(brush, 0, 0, width, height);
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }
    }

    // Creates a Word document containing the specified JPEG image
    private static void CreateDocumentWithJpeg(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document containing image: {Path.GetFileName(imagePath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Applies a simple vignette effect to the bitmap using a radial gradient
    private static void ApplyVignette(Bitmap bitmap)
    {
        int w = bitmap.Width;
        int h = bitmap.Height;

        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Create an elliptical path covering the whole image
            using (GraphicsPath path = new GraphicsPath())
            {
                path.AddEllipse(0, 0, w, h);
                using (PathGradientBrush brush = new PathGradientBrush(path))
                {
                    brush.CenterColor = Color.FromArgb(0, 0, 0, 0); // Transparent center
                    brush.SurroundColors = new Color[] { Color.FromArgb(180, 0, 0, 0) }; // Semi‑transparent dark edges
                    g.FillRectangle(brush, 0, 0, w, h);
                }
            }
        }
    }
}
