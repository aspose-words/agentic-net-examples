using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (sample.png).
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.LightBlue);
        // Draw a simple rectangle for visual distinction.
        using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5))
        {
            graphics.DrawRectangle(pen, 20, 20, 160, 160);
        }
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a Word document and insert the sample image twice.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.InsertParagraph(); // separate images
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        List<string> extractedImageFiles = new List<string>();

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(artifactsDir, imageFileName);
                shape.ImageData.Save(imageFullPath);
                extractedImageFiles.Add(imageFileName);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // 4. Generate a LaTeX file referencing each extracted image.
        // -------------------------------------------------
        string texPath = Path.Combine(artifactsDir, "output.tex");
        using (StreamWriter writer = new StreamWriter(texPath))
        {
            writer.WriteLine("\\documentclass{article}");
            writer.WriteLine("\\usepackage{graphicx}");
            writer.WriteLine("\\begin{document}");
            foreach (string imgFile in extractedImageFiles)
            {
                writer.WriteLine($"\\includegraphics[width=\\linewidth]{{{imgFile}}}");
                writer.WriteLine("\\\\"); // line break between images
            }
            writer.WriteLine("\\end{document}");
        }

        // Optional: inform the user where files are located.
        Console.WriteLine($"Artifacts written to: {artifactsDir}");
    }
}
