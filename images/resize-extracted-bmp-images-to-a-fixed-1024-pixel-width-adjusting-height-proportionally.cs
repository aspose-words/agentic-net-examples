using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create sample BMP images.
        string[] sampleImagePaths = { "sample1.bmp", "sample2.bmp" };
        CreateSampleBmp(sampleImagePaths[0], 800, 600, Aspose.Drawing.Color.LightBlue);
        CreateSampleBmp(sampleImagePaths[1], 1200, 900, Aspose.Drawing.Color.LightCoral);

        // Step 2: Insert images into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string imgPath in sampleImagePaths)
        {
            builder.InsertImage(imgPath);
            builder.Writeln(); // separate images
        }
        string docPath = "sample.docx";
        doc.Save(docPath);

        // Step 3: Extract BMP images from the document and resize them.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        int resizedCount = 0;
        int index = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Save extracted image directly to a BMP file.
            string extractedPath = $"extracted-{index}.bmp";
            shape.ImageData.Save(extractedPath);
            if (!File.Exists(extractedPath))
                throw new Exception($"Failed to save extracted image {extractedPath}.");

            extractedCount++;

            // Resize extracted BMP to width 1024 while preserving aspect ratio.
            using (Bitmap originalBmp = new Bitmap(extractedPath))
            {
                int originalWidth = originalBmp.Width;
                int originalHeight = originalBmp.Height;
                int newWidth = 1024;
                int newHeight = (int)Math.Round(originalHeight * (newWidth / (double)originalWidth));

                using (Bitmap resizedBmp = new Bitmap(newWidth, newHeight))
                {
                    using (Graphics g = Graphics.FromImage(resizedBmp))
                    {
                        g.Clear(Aspose.Drawing.Color.White);
                        g.DrawImage(originalBmp, new Rectangle(0, 0, newWidth, newHeight));
                    }

                    string resizedPath = $"resized-{index}.bmp";
                    resizedBmp.Save(resizedPath);
                    if (!File.Exists(resizedPath))
                        throw new Exception($"Failed to save resized image {resizedPath}.");

                    resizedCount++;
                }
            }

            index++;
        }

        // Validation.
        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");
        if (resizedCount == 0)
            throw new Exception("No images were resized.");

        Console.WriteLine($"Extraction complete: {extractedCount} image(s) extracted.");
        Console.WriteLine($"Resizing complete: {resizedCount} image(s) resized to width 1024.");
    }

    private static void CreateSampleBmp(string path, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bmp = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(backColor);
                // Draw a simple rectangle for visual distinction.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bmp.Save(path);
        }

        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image {path}.");
    }
}
