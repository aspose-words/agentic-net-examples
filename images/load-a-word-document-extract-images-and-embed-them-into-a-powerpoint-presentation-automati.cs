using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image.
        string imagePath = "input.png";
        CreateSampleImage(imagePath);

        // Create a Word document that contains the sample image.
        string wordPath = "sample.docx";
        CreateWordDocumentWithImage(wordPath, imagePath);

        // Extract all images from the Word document.
        List<string> extractedImages = ExtractImagesFromWord(wordPath);

        // NOTE: PowerPoint creation is omitted because Aspose.Slides is not available
        // in the required package set. The extracted images are saved to disk and can
        // be used in a presentation by external tools.

        // Validate that all expected files were created.
        ValidateFileExists(imagePath);
        ValidateFileExists(wordPath);
        foreach (string img in extractedImages)
        {
            ValidateFileExists(img);
        }
    }

    // Creates a simple 200x200 PNG image with a blue rectangle.
    private static void CreateSampleImage(string path)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);
        Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5);
        g.DrawRectangle(pen, 20, 20, 160, 160);
        pen.Dispose();
        g.Dispose();
        bitmap.Save(path);
        bitmap.Dispose();
    }

    // Creates a Word document and inserts the previously created image.
    private static void CreateWordDocumentWithImage(string docPath, string imgPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imgPath);
        doc.Save(docPath);
    }

    // Extracts all images from the Word document and saves them as separate files.
    private static List<string> ExtractImagesFromWord(string docPath)
    {
        List<string> extracted = new List<string>();
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int index = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string outPath = $"extracted-{index}.png";
                shape.ImageData.Save(outPath);
                extracted.Add(outPath);
                index++;
            }
        }
        if (extracted.Count == 0)
        {
            throw new Exception("No images were extracted from the Word document.");
        }
        return extracted;
    }

    // Validates that a file exists at the given path.
    private static void ValidateFileExists(string path)
    {
        if (!File.Exists(path))
        {
            throw new Exception($"Expected file not found: {path}");
        }
    }
}
