using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- Create a sample image ----------
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.LightBlue);
        // Additional deterministic drawing can be added here if desired
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // ---------- Create a Word document and insert the image ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.Writeln();
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(outputDir, "sample.docx");
        doc.Save(docPath);

        // ---------- Load the document ----------
        Document loadedDoc = new Document(docPath);

        // ---------- Extract images ----------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        List<string> latexLines = new List<string>();

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(outputDir, imageFileName);
                shape.ImageData.Save(imageFullPath);
                latexLines.Add($"\\includegraphics{{{imageFileName}}}");
                imageIndex++;
            }
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // ---------- Generate LaTeX file ----------
        string texPath = Path.Combine(outputDir, "document.tex");
        using (StreamWriter writer = new StreamWriter(texPath))
        {
            writer.WriteLine("\\documentclass{article}");
            writer.WriteLine("\\usepackage{graphicx}");
            writer.WriteLine("\\begin{document}");
            foreach (string line in latexLines)
                writer.WriteLine(line);
            writer.WriteLine("\\end{document}");
        }

        // Optional: indicate completion
        Console.WriteLine($"Extraction complete. Files are located in: {outputDir}");
    }
}
