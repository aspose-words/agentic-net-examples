using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image (input.png).
        const string sampleImagePath = "input.png";
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // Additional drawing can be added here if desired.
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Step 2: Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Step 3: Load the document for image extraction.
        Document loadedDoc = new Document(docPath);

        // Step 4: Extract images from the document.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> extractedImageFiles = new List<string>();
        int imageIndex = 1;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extension = GetExtension(shape.ImageData.ImageType);
                string imageFileName = $"image{imageIndex}{extension}";
                shape.ImageData.Save(imageFileName);
                extractedImageFiles.Add(imageFileName);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Step 5: Generate a LaTeX file referencing each extracted image.
        const string latexPath = "output.tex";
        StringBuilder latexBuilder = new StringBuilder();
        latexBuilder.AppendLine("\\documentclass{article}");
        latexBuilder.AppendLine("\\usepackage{graphicx}");
        latexBuilder.AppendLine("\\begin{document}");

        foreach (string imgFile in extractedImageFiles)
        {
            latexBuilder.AppendLine("\\begin{figure}[h]");
            latexBuilder.AppendLine("\\centering");
            latexBuilder.AppendLine($"\\includegraphics[width=\\linewidth]{{{imgFile}}}");
            latexBuilder.AppendLine("\\caption{Extracted image}");
            latexBuilder.AppendLine("\\end{figure}");
        }

        latexBuilder.AppendLine("\\end{document}");
        File.WriteAllText(latexPath, latexBuilder.ToString());

        // Validate that the LaTeX file was created.
        if (!File.Exists(latexPath))
            throw new InvalidOperationException("Failed to create the LaTeX file.");
    }

    // Helper method to map Aspose.Words.Drawing.ImageType to a file extension.
    private static string GetExtension(ImageType imageType)
    {
        switch (imageType)
        {
            case ImageType.Jpeg:
                return ".jpg";
            case ImageType.Png:
                return ".png";
            case ImageType.Gif:
                return ".gif";
            case ImageType.Bmp:
                return ".bmp";
            case ImageType.Emf:
                return ".emf";
            case ImageType.Wmf:
                return ".wmf";
            default:
                return ".img";
        }
    }
}
