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
        // Create a deterministic sample image.
        const string sampleImagePath = "input.png";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle for visual distinction.
                graphics.DrawRectangle(new Pen(Color.Black), 20, 20, 160, 160);
            }
            bitmap.Save(sampleImagePath);
        }

        // Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Load the document for image extraction.
        Document loadedDoc = new Document(docPath);

        // Extract images from the document.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> extractedImageFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string imageFileName = $"image{imageIndex}.png";
                shape.ImageData.Save(imageFileName);
                extractedImageFiles.Add(imageFileName);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Create a Markdown file with image links.
        const string markdownPath = "output.md";
        using (StreamWriter writer = new StreamWriter(markdownPath))
        {
            foreach (string imgFile in extractedImageFiles)
            {
                writer.WriteLine($"![]({imgFile})");
            }
        }

        // Validate that the markdown file was created.
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("Markdown file was not created.");

        // Optional: Clean up temporary files (comment out if you want to keep them).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
