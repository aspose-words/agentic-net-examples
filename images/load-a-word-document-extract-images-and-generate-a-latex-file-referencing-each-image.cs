using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image.
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath);

        // Step 2: Create a Word document and insert the sample image twice.
        const string docPath = "sample.docx";
        CreateWordDocumentWithImages(docPath, sampleImagePath);

        // Step 3: Load the document and extract all images.
        Document doc = new Document(docPath);
        List<string> extractedImagePaths = ExtractImages(doc);

        // Validate that at least one image was extracted.
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Step 4: Generate a LaTeX file referencing each extracted image.
        const string latexPath = "output.tex";
        GenerateLatexFile(latexPath, extractedImagePaths);

        // Validate LaTeX file creation.
        if (!File.Exists(latexPath))
            throw new InvalidOperationException("LaTeX file was not created.");

        // Example completed.
    }

    private static void CreateSampleImage(string path)
    {
        // Create a 200x200 white bitmap.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple black rectangle.
                using (Pen pen = new Pen(Color.Black, 3))
                {
                    graphics.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            // Save the bitmap as PNG.
            bitmap.Save(path, ImageFormat.Png);
        }

        // Ensure the image file exists.
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample image at '{path}'.");
    }

    private static void CreateWordDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image twice, each on its own paragraph.
        builder.Writeln("First image:");
        builder.InsertImage(imagePath);
        builder.Writeln();
        builder.Writeln("Second image:");
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);

        // Validate document creation.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create Word document at '{docPath}'.");
    }

    private static List<string> ExtractImages(Document doc)
    {
        List<string> imagePaths = new List<string>();
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int index = 1;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string imageFileName = $"extracted-{index}.png";
                shape.ImageData.Save(imageFileName);
                if (!File.Exists(imageFileName))
                    throw new InvalidOperationException($"Failed to save extracted image '{imageFileName}'.");
                imagePaths.Add(imageFileName);
                index++;
            }
        }

        return imagePaths;
    }

    private static void GenerateLatexFile(string latexPath, List<string> imagePaths)
    {
        using (StreamWriter writer = new StreamWriter(latexPath, false))
        {
            writer.WriteLine(@"\documentclass{article}");
            writer.WriteLine(@"\usepackage{graphicx}");
            writer.WriteLine(@"\begin{document}");
            writer.WriteLine(@"\section*{Extracted Images}");

            int imgIndex = 1;
            foreach (string imgPath in imagePaths)
            {
                writer.WriteLine(@"\begin{figure}[h]");
                writer.WriteLine(@"\centering");
                writer.WriteLine($@"\includegraphics[width=0.8\textwidth]{{{imgPath}}}");
                writer.WriteLine($@"\caption{{Image {imgIndex}}}");
                writer.WriteLine(@"\end{figure}");
                writer.WriteLine();
                imgIndex++;
            }

            writer.WriteLine(@"\end{document}");
        }
    }
}
