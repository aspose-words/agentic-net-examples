using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (sample.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the sample image twice.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateWordDocumentWithImages(docPath, sampleImagePath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images.
        // -----------------------------------------------------------------
        List<string> extractedImagePaths = ExtractImagesFromDocument(docPath, artifactsDir);

        // Validate that at least one image was extracted.
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. Generate a LaTeX file that references each extracted image.
        // -----------------------------------------------------------------
        string texPath = Path.Combine(artifactsDir, "output.tex");
        GenerateLatexFile(texPath, extractedImagePaths);

        // The example finishes execution here.
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create a bitmap, fill it with white, and save it.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Additional deterministic drawing can be added here if needed.
            }
            bitmap.Save(filePath);
        }
    }

    private static void CreateWordDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image twice, separated by a paragraph.
        builder.InsertImage(imagePath);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);

        doc.Save(docPath);
    }

    private static List<string> ExtractImagesFromDocument(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        var extractedPaths = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(outputDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imageFullPath);
                extractedPaths.Add(imageFullPath);
                imageIndex++;
            }
        }

        return extractedPaths;
    }

    private static void GenerateLatexFile(string texFilePath, List<string> imagePaths)
    {
        using (StreamWriter writer = new StreamWriter(texFilePath))
        {
            writer.WriteLine(@"\documentclass{article}");
            writer.WriteLine(@"\usepackage{graphicx}");
            writer.WriteLine(@"\begin{document}");
            writer.WriteLine(@"\section*{Extracted Images}");

            for (int i = 0; i < imagePaths.Count; i++)
            {
                string fileName = Path.GetFileName(imagePaths[i]);
                writer.WriteLine(@"\begin{figure}[h]");
                writer.WriteLine(@"\centering");
                writer.WriteLine($@"\includegraphics[width=\linewidth]{{{fileName}}}");
                writer.WriteLine($@"\caption{{Image {i}}}");
                writer.WriteLine(@"\end{figure}");
                writer.WriteLine();
            }

            writer.WriteLine(@"\end{document}");
        }
    }
}
