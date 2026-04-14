using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string imagesDir = Path.Combine(artifactsDir, "Images");
        string docsDir = Path.Combine(artifactsDir, "Docs");
        string outputDir = Path.Combine(artifactsDir, "WebP");

        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create deterministic sample images (PNG) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string samplePng1 = Path.Combine(imagesDir, "sample1.png");
        Aspose.Drawing.Bitmap bitmap1 = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics g1 = Aspose.Drawing.Graphics.FromImage(bitmap1);
        g1.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
        {
            g1.DrawRectangle(pen, 20, 20, 160, 160);
        }
        bitmap1.Save(samplePng1);
        g1.Dispose();
        bitmap1.Dispose();

        string samplePng2 = Path.Combine(imagesDir, "sample2.png");
        Aspose.Drawing.Bitmap bitmap2 = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics g2 = Aspose.Drawing.Graphics.FromImage(bitmap2);
        g2.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
        {
            g2.DrawEllipse(pen, 20, 20, 160, 160);
        }
        bitmap2.Save(samplePng2);
        g2.Dispose();
        bitmap2.Dispose();

        // ---------------------------------------------------------------
        // 2. Create sample Word documents and insert the images.
        // ---------------------------------------------------------------
        string docPath1 = Path.Combine(docsDir, "doc1.docx");
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.InsertImage(samplePng1);
        builder1.Writeln();
        builder1.InsertImage(samplePng2);
        doc1.Save(docPath1);

        string docPath2 = Path.Combine(docsDir, "doc2.docx");
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.InsertImage(samplePng2);
        builder2.Writeln();
        builder2.InsertImage(samplePng1);
        doc2.Save(docPath2);

        // ---------------------------------------------------------------
        // 3. Batch process all Word files: extract each image and convert to WebP.
        // ---------------------------------------------------------------
        int totalConverted = 0;
        string[] wordFiles = Directory.GetFiles(docsDir, "*.docx");

        foreach (string wordFile in wordFiles)
        {
            Document srcDoc = new Document(wordFile);
            NodeCollection shapeNodes = srcDoc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Extract the image bytes into a memory stream.
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0; // Reset before reuse.

                    // Create a temporary document that contains only this image.
                    Document tempDoc = new Document();
                    DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                    tempBuilder.InsertImage(imageStream);

                    // Save the temporary document as a WebP image.
                    string outFileName = $"img_{Path.GetFileNameWithoutExtension(wordFile)}_{imageIndex}.webp";
                    string outPath = Path.Combine(outputDir, outFileName);
                    ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP);
                    tempDoc.Save(outPath, webpOptions);

                    // Validate that the WebP file was created.
                    if (!File.Exists(outPath))
                        throw new InvalidOperationException($"Failed to create WebP file: {outPath}");

                    totalConverted++;
                }

                imageIndex++;
            }
        }

        // Ensure that at least one image was converted.
        if (totalConverted == 0)
            throw new InvalidOperationException("No images were extracted and converted to WebP.");
    }
}
