using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names
        const string imageFile = "sample.png";
        const string docFile = "sample.docm";

        // Create a sample PNG image
        CreateSamplePng(imageFile, 100, 100);

        // Create a DOCM document and insert the image as shapes with names
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first image and set shape name
        Shape shape1 = builder.InsertImage(imageFile);
        shape1.Name = "FirstShape";

        // Insert second image (reuse same PNG) and set a different shape name
        Shape shape2 = builder.InsertImage(imageFile);
        shape2.Name = "SecondShape";

        // Save the document as DOCM
        doc.Save(docFile, SaveFormat.Docm);

        // Load the DOCM document for extraction
        Document loadedDoc = new Document(docFile);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Node node in shapeNodes)
        {
            Shape shape = (Shape)node;
            if (!shape.HasImage)
                continue;

            // Determine file extension based on image type
            string extension = GetExtensionFromImageType(shape.ImageData.ImageType);
            if (string.IsNullOrEmpty(extension))
                continue; // Skip unsupported types

            // Use the shape's name for the output file; fallback to a generic name if missing
            string baseName = !string.IsNullOrEmpty(shape.Name) ? shape.Name : $"shape_{extractedCount + 1}";
            string outputFileName = $"{baseName}{extension}";

            // Save the image
            shape.ImageData.Save(outputFileName);
            extractedCount++;
        }

        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        // Optional clean‑up (commented out to keep the generated files for inspection)
        // File.Delete(imageFile);
        // File.Delete(docFile);
    }

    private static void CreateSamplePng(string filePath, int width, int height)
    {
        // Create a deterministic PNG image using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Red, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    private static string GetExtensionFromImageType(ImageType imageType)
    {
        // Map Aspose.Words.Drawing.ImageType values to file extensions
        switch (imageType)
        {
            case ImageType.Jpeg:
                return ".jpeg";
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
                return string.Empty; // Unsupported or unknown type
        }
    }
}
