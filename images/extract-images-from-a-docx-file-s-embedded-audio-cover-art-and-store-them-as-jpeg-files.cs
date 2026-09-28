using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names
        string workDir = Directory.GetCurrentDirectory();
        string imagePath = Path.Combine(workDir, "cover.png");
        string docPath = Path.Combine(workDir, "sample.docx");
        string outputDir = Path.Combine(workDir, "extracted");
        Directory.CreateDirectory(outputDir);

        // Create a sample image (cover art)
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.LightBlue))
                {
                    g.FillRectangle(brush, 20, 20, imgWidth - 40, imgHeight - 40);
                }
            }
            bitmap.Save(imagePath);
        }

        // Create a DOCX and insert the image (simulating audio cover art)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // Load the document and extract all images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outFile = Path.Combine(outputDir, $"extracted-{++extractedCount}.png");
                shape.ImageData.Save(outFile);
            }
        }

        // Validate that at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Summary output
        Console.WriteLine($"Created sample document: {docPath}");
        Console.WriteLine($"Extracted {extractedCount} image(s) to folder: {outputDir}");
    }
}
