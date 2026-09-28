using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = "InputDocs";
        string outputFolder = "OutputImages";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample JPEG image
        string sampleImagePath = Path.Combine(inputFolder, "sample.jpg");
        CreateSampleJpeg(sampleImagePath);

        // Create sample Word documents that contain the JPEG image
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleImagePath);

        // Process each document: extract JPEG images, apply vignette, save result
        int totalProcessed = 0;
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage)
                    continue;

                if (shape.ImageData.ImageType != ImageType.Jpeg)
                    continue;

                // Extract image to memory stream
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0;

                    // Load image into bitmap
                    using (Bitmap bitmap = new Bitmap(imgStream))
                    {
                        ApplyVignetteEffect(bitmap);

                        // Save processed image
                        string outputPath = Path.Combine(
                            outputFolder,
                            $"vignette-{Path.GetFileNameWithoutExtension(docPath)}-{imageIndex}.jpg");
                        bitmap.Save(outputPath, ImageFormat.Jpeg);
                        totalProcessed++;
                    }
                }

                imageIndex++;
            }
        }

        // Validation
        if (totalProcessed == 0)
            throw new InvalidOperationException("No JPEG images were processed.");

        // Example completed
    }

    private static void CreateSampleJpeg(string path)
    {
        int width = 300;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.FromArgb(255, 100, 150, 200)))
                {
                    g.FillRectangle(brush, 0, 0, width, height);
                }
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 5))
                {
                    g.DrawEllipse(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }

    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static void ApplyVignetteEffect(Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;
        Rectangle rect = new Rectangle(0, 0, width, height);

        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Create a radial gradient brush (transparent center, dark edges)
            using (GraphicsPath path = new GraphicsPath())
            {
                path.AddEllipse(rect);
                using (PathGradientBrush brush = new PathGradientBrush(path))
                {
                    brush.CenterColor = Aspose.Drawing.Color.FromArgb(0, 0, 0, 0); // fully transparent
                    brush.SurroundColors = new Aspose.Drawing.Color[] {
                        Aspose.Drawing.Color.FromArgb(180, 0, 0, 0) // semi‑transparent black
                    };
                    graphics.FillRectangle(brush, rect);
                }
            }
        }
    }
}
