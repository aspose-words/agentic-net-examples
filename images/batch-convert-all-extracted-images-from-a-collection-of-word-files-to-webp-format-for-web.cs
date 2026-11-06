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
        string inputDir = "InputDocs";
        string outputDir = "OutputImages";
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample image (sample.png)
        string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath);

        // Create sample Word documents with the image inserted
        CreateSampleWordDocument(Path.Combine(inputDir, "Doc1.docx"), sampleImagePath);
        CreateSampleWordDocument(Path.Combine(inputDir, "Doc2.docx"), sampleImagePath);

        // Batch process: extract images and convert to WebP
        int extractedCount = 0;
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    // Get image bytes
                    byte[] imageBytes = shape.ImageData.ImageBytes;

                    // Load image into Aspose.Drawing.Image
                    using (MemoryStream ms = new MemoryStream(imageBytes))
                    {
                        ms.Position = 0;
                        using (Image img = Image.FromStream(ms))
                        {
                            // Build output file name
                            string docName = Path.GetFileNameWithoutExtension(docPath);
                            string outputFile = Path.Combine(outputDir, $"{docName}_img{imageIndex}.webp");

                            // Save as WebP (format inferred from extension)
                            img.Save(outputFile);
                            extractedCount++;
                        }
                    }
                    imageIndex++;
                }
            }
        }

        // Validation
        if (extractedCount == 0)
        {
            throw new Exception("No images were extracted and converted.");
        }

        // Clean up sample image file
        if (File.Exists(sampleImagePath))
        {
            File.Delete(sampleImagePath);
        }
    }

    private static void CreateSampleImage(string path)
    {
        int width = 200;
        int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path);
        }
    }

    private static void CreateSampleWordDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
