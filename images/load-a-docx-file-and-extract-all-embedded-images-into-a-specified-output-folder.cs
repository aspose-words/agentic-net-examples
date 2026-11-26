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
        // Paths
        string inputDocPath = "sample.docx";
        string sampleImagePath = "sample.png";
        string outputFolder = "ExtractedImages";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // Create a deterministic sample image (if not exists)
        // -------------------------------------------------
        if (!File.Exists(sampleImagePath))
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
                }
                bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
            }
        }

        // -------------------------------------------------
        // Create a sample DOCX containing the image (if not exists)
        // -------------------------------------------------
        if (!File.Exists(inputDocPath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Document with an embedded image:");
            builder.InsertImage(sampleImagePath);
            doc.Save(inputDocPath);
        }

        // -------------------------------------------------
        // Load the DOCX file
        // -------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);

        // -------------------------------------------------
        // Extract all embedded images
        // -------------------------------------------------
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                // Determine file extension from the image type
                string extension = shape.ImageData.ImageType.ToString().ToLowerInvariant();
                string outputPath = Path.Combine(outputFolder, $"Image-{imageIndex}.{extension}");
                shape.ImageData.Save(outputPath);
                imageIndex++;
            }
        }

        // -------------------------------------------------
        // Validation
        // -------------------------------------------------
        if (imageIndex == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        Console.WriteLine($"Extraction complete. {imageIndex} image(s) saved to \"{outputFolder}\".");
    }
}
