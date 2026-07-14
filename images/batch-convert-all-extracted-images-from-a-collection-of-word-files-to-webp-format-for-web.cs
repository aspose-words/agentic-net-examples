using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageConversion
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string outputWebpDir = Path.Combine(baseDir, "OutputWebP");
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(outputWebpDir);

        // Create sample images (PNG and JPEG) that will be inserted into the Word files.
        string pngPath = Path.Combine(baseDir, "sample1.png");
        string jpgPath = Path.Combine(baseDir, "sample2.jpg");
        CreateSamplePng(pngPath);
        CreateSampleJpeg(jpgPath);

        // Create a couple of Word documents each containing the sample images.
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc1.docx"), pngPath, jpgPath);
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc2.docx"), jpgPath, pngPath);

        // Process each Word document in the input folder.
        var docFiles = Directory.GetFiles(inputDocsDir, "*.docx");
        int totalConverted = 0;

        foreach (var docFile in docFiles)
        {
            // Load the document.
            Document doc = new Document(docFile);

            // Get all shapes that contain images.
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage)
                            .ToList();

            int shapeIndex = 0;
            foreach (var shape in shapes)
            {
                // Extract the image bytes from the shape.
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0;

                    // Insert the extracted image into a temporary one‑page document.
                    Document tempDoc = new Document();
                    DocumentBuilder builder = new DocumentBuilder(tempDoc);
                    builder.InsertImage(imageStream);

                    // Define the output WebP file name.
                    string outputFileName = $"Img_{Path.GetFileNameWithoutExtension(docFile)}_{shapeIndex}.webp";
                    string outputPath = Path.Combine(outputWebpDir, outputFileName);

                    // Save the temporary document as a WebP image.
                    ImageSaveOptions options = new ImageSaveOptions(SaveFormat.WebP);
                    tempDoc.Save(outputPath, options);

                    // Validate that the file was created.
                    if (!File.Exists(outputPath))
                        throw new InvalidOperationException($"Failed to create WebP file: {outputPath}");

                    totalConverted++;
                }

                shapeIndex++;
            }
        }

        // Ensure that at least one image was converted.
        if (totalConverted == 0)
            throw new InvalidOperationException("No images were found and converted.");

        Console.WriteLine($"Converted {totalConverted} image(s) to WebP format in '{outputWebpDir}'.");
    }

    // Creates a deterministic PNG image.
    private static void CreateSamplePng(string filePath)
    {
        int width = 200;
        int height = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
                {
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                    {
                        g.DrawString("PNG Sample", font, brush, new Aspose.Drawing.PointF(20, 40));
                    }
                }
            }
            bitmap.Save(filePath);
        }
    }

    // Creates a deterministic JPEG image.
    private static void CreateSampleJpeg(string filePath)
    {
        int width = 200;
        int height = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Green, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
                {
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                    {
                        g.DrawString("JPEG Sample", font, brush, new Aspose.Drawing.PointF(20, 40));
                    }
                }
            }
            // Save as JPEG using the Aspose.Drawing.Imaging.ImageFormat enumeration.
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }
    }

    // Creates a Word document containing two images.
    private static void CreateSampleDocument(string docPath, string firstImagePath, string secondImagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First image:");
        builder.InsertImage(firstImagePath);
        builder.Writeln();
        builder.Writeln("Second image:");
        builder.InsertImage(secondImagePath);

        doc.Save(docPath);
    }
}
