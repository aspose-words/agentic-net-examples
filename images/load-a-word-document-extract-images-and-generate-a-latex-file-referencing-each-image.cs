using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // For Bitmap, Graphics, Color

public class Program
{
    public static void Main()
    {
        // Define deterministic file names.
        const string sampleImagePath = "sample.png";
        const string docPath = "sample.docx";
        const string latexPath = "output.tex";

        // -----------------------------------------------------------------
        // 1. Create a sample image (100x100 white bitmap) using Aspose.Drawing.
        // -----------------------------------------------------------------
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // (Optional) draw something deterministic here if desired.
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the sample image twice.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.InsertParagraph(); // separate the images
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        List<Shape> shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                          .Cast<Shape>()
                                          .Where(s => s.HasImage)
                                          .ToList();

        if (!shapeNodes.Any())
            throw new InvalidOperationException("No images were found in the document.");

        List<string> extractedImageFiles = new List<string>();
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string imageFileName = $"image_{imageIndex}{extension}";
            shape.ImageData.Save(imageFileName);
            extractedImageFiles.Add(imageFileName);
            imageIndex++;
        }

        // -----------------------------------------------------------------
        // 4. Generate a simple LaTeX file that includes each extracted image.
        // -----------------------------------------------------------------
        using (StreamWriter writer = new StreamWriter(latexPath, false))
        {
            writer.WriteLine(@"\documentclass{article}");
            writer.WriteLine(@"\usepackage{graphicx}");
            writer.WriteLine(@"\begin{document}");
            writer.WriteLine();

            for (int i = 0; i < extractedImageFiles.Count; i++)
            {
                string imgFile = extractedImageFiles[i];
                writer.WriteLine(@"\begin{figure}[h]");
                writer.WriteLine(@"\centering");
                writer.WriteLine($@"\includegraphics[width=0.8\textwidth]{{{imgFile}}}");
                writer.WriteLine($@"\caption{{Image {i}}}");
                writer.WriteLine(@"\end{figure}");
                writer.WriteLine();
            }

            writer.WriteLine(@"\end{document}");
        }

        // -----------------------------------------------------------------
        // 5. Validation: ensure LaTeX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(latexPath))
            throw new InvalidOperationException("LaTeX file was not created.");

        // Program completed successfully.
    }
}
