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
        // Paths for temporary files
        const string sampleImagePath = "sample.png";
        const string docPath = "sample.docx";
        const string latexPath = "output.tex";

        // 1. Create a deterministic sample image using Aspose.Drawing
        const int imgWidth = 200;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white
                g.Clear(Color.White);
                // Draw a simple black rectangle
                g.DrawRectangle(Pens.Black, 10, 10, imgWidth - 20, imgHeight - 20);
            }
            // Save the image to disk
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // 2. Create a Word document and insert the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // 3. Load the Word document (simulating a separate load operation)
        Document loadedDoc = new Document(docPath);

        // 4. Extract images from the document
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> extractedImageFiles = new List<string>();
        int imageIndex = 1;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string imageFileName = $"image_{imageIndex}.png";
                shape.ImageData.Save(imageFileName);
                extractedImageFiles.Add(imageFileName);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // 5. Generate LaTeX file referencing each extracted image
        using (StreamWriter writer = new StreamWriter(latexPath, false))
        {
            writer.WriteLine(@"\documentclass{article}");
            writer.WriteLine(@"\usepackage{graphicx}");
            writer.WriteLine(@"\begin{document}");
            foreach (string imgFile in extractedImageFiles)
            {
                writer.WriteLine($@"\includegraphics{{{imgFile}}}");
                writer.WriteLine(); // Add a blank line for readability
            }
            writer.WriteLine(@"\end{document}");
        }

        // Validate LaTeX file creation
        if (!File.Exists(latexPath))
            throw new InvalidOperationException("LaTeX file was not created.");

        // Optional: Clean up temporary files (comment out if you want to inspect them)
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
        // foreach (var img in extractedImageFiles) File.Delete(img);
        // File.Delete(latexPath);
    }
}
