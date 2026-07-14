using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample bitmap that represents an equation.
        const string equationImagePath = "equation.png";
        const int width = 200;
        const int height = 80;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw simple equation text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    g.DrawString("a + b = c", font, Aspose.Drawing.Brushes.Black, new PointF(10, 20));
                }
            }
            bitmap.Save(equationImagePath, ImageFormat.Png);
        }

        // Create a Word document and insert the equation image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with an equation image:");
        builder.InsertImage(equationImagePath);
        doc.Save(docPath);

        // Load the document and extract images from shape nodes.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string outputImagePath = $"extracted-{extractedCount}.png";
                shape.ImageData.Save(outputImagePath);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Clean up temporary files (optional).
        // File.Delete(equationImagePath);
        // File.Delete(docPath);
    }
}
