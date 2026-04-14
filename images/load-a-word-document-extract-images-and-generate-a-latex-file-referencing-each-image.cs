using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageExtractionToLatex
{
    public static void Main()
    {
        // Prepare deterministic file names
        const string imageFile = "input.png";
        const string docFile = "sample.docx";
        const string latexFile = "output.tex";

        // -----------------------------------------------------------------
        // 1. Create a sample image using Aspose.Drawing and save it locally
        // -----------------------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle for visual distinction
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(imageFile, ImageFormat.Png);
        }

        // ---------------------------------------------------------------
        // 2. Create a Word document and insert the sample image twice
        // ---------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imageFile);
        builder.Writeln(); // separate images with a line break
        builder.InsertImage(imageFile);
        doc.Save(docFile);

        // ---------------------------------------------------------------
        // 3. Load the document and extract all images from shape nodes
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docFile);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> extractedImageFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine proper file extension based on image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedFileName = $"extracted_{imageIndex}{extension}";
                shape.ImageData.Save(extractedFileName);
                extractedImageFiles.Add(extractedFileName);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // ---------------------------------------------------------------
        // 4. Generate a LaTeX file that references each extracted image
        // ---------------------------------------------------------------
        using (StreamWriter writer = new StreamWriter(latexFile, false))
        {
            writer.WriteLine(@"\documentclass{article}");
            writer.WriteLine(@"\usepackage{graphicx}");
            writer.WriteLine(@"\begin{document}");
            foreach (string imgPath in extractedImageFiles)
            {
                writer.WriteLine($@"\begin{{figure}}[h]");
                writer.WriteLine($@"\centering");
                writer.WriteLine($@"\includegraphics[width=0.5\textwidth]{{{imgPath}}}");
                writer.WriteLine($@"\caption{{Extracted image {Path.GetFileNameWithoutExtension(imgPath)}}}");
                writer.WriteLine($@"\end{{figure}}");
                writer.WriteLine();
            }
            writer.WriteLine(@"\end{document}");
        }

        // Validate that the LaTeX file was created
        if (!File.Exists(latexFile))
            throw new InvalidOperationException("Failed to create the LaTeX output file.");

        // Optional: clean up resources (files remain for inspection)
        Console.WriteLine("Image extraction and LaTeX generation completed successfully.");
    }
}
