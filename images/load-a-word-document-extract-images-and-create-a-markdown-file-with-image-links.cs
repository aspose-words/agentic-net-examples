using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // (Optional) draw something simple.
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Build a Word document and insert the sample image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an image:");
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document back.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // 4. Extract all images from the document.
        // -----------------------------------------------------------------
        string imagesFolder = Path.Combine(artifactsDir, "Images");
        Directory.CreateDirectory(imagesFolder);

        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> extractedImageFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesFolder, imageFileName);
                shape.ImageData.Save(imageFullPath);
                extractedImageFiles.Add(imageFileName);
                imageIndex++;
            }
        }

        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 5. Create a Markdown file with links to the extracted images.
        // -----------------------------------------------------------------
        string markdownPath = Path.Combine(artifactsDir, "output.md");
        using (StreamWriter writer = new StreamWriter(markdownPath))
        {
            writer.WriteLine("# Extracted Images");
            writer.WriteLine();

            for (int i = 0; i < extractedImageFiles.Count; i++)
            {
                // Use forward slashes for Markdown compatibility.
                string relativePath = $"Images/{extractedImageFiles[i]}";
                writer.WriteLine($"![Image {i}]({relativePath})");
                writer.WriteLine();
            }
        }

        // -----------------------------------------------------------------
        // 6. Validate that the Markdown file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(markdownPath))
            throw new FileNotFoundException("Markdown file was not created.", markdownPath);
    }
}
