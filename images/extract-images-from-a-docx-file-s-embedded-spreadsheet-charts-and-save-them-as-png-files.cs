using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample chart image (PNG)
        const string chartImagePath = "chart.png";
        const int chartWidth = 400;
        const int chartHeight = 300;

        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(chartWidth, chartHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 50, 50, 300, 200);
                }
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24, Aspose.Drawing.FontStyle.Bold))
                {
                    g.DrawString("Sample Chart", font, new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black), new Aspose.Drawing.PointF(80, 130));
                }
            }
            bitmap.Save(chartImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // Create a DOCX document and embed the chart image
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with an embedded chart image:");
        builder.InsertImage(chartImagePath);
        doc.Save(docPath);

        // Load the document and extract images from shapes
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outputImagePath = $"extracted-{extractedCount + 1}.png";
                shape.ImageData.Save(outputImagePath);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Optional cleanup (commented out)
        // File.Delete(chartImagePath);
        // File.Delete(docPath);
    }
}
