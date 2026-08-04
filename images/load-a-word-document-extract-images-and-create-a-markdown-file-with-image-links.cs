using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string imagesDir = Path.Combine(artifactsDir, "Images");
        Directory.CreateDirectory(imagesDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (sample.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                // Simple drawing – a red ellipse.
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawEllipse(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the sample image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "SampleDocument.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Below is a sample image inserted into the document:");
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document (demonstrating load via file name).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // 4. Extract all images from the document.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        var extractedImageFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);
                shape.ImageData.Save(imageFullPath);
                extractedImageFiles.Add(imageFileName);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 5. Generate a Markdown file with links to the extracted images.
        // -----------------------------------------------------------------
        string markdownPath = Path.Combine(artifactsDir, "DocumentImages.md");
        using (StreamWriter writer = new StreamWriter(markdownPath, false))
        {
            writer.WriteLine("# Extracted Images");
            writer.WriteLine();

            foreach (string imgFile in extractedImageFiles)
            {
                // Use a relative path to the Images folder.
                string relativePath = Path.Combine("Images", imgFile).Replace('\\', '/');
                writer.WriteLine($"![]({relativePath})");
                writer.WriteLine();
            }
        }

        // -----------------------------------------------------------------
        // 6. Simple verification that the Markdown file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(markdownPath))
            throw new FileNotFoundException("Markdown file was not created.", markdownPath);
    }
}
