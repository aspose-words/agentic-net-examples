using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names
        string imagePath = "sample.png";
        string docPath = "sample.docx";
        string outputFolder = "extracted_images";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Create a sample PNG image using Aspose.Drawing
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                    {
                        g.DrawString("Sample", font, brush, new Aspose.Drawing.PointF(40, 80));
                    }
                }
            }
            bitmap.Save(imagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // Create a new Word document and insert the sample image and a chart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the sample image
        builder.InsertImage(imagePath);

        // Insert a chart (embedded spreadsheet chart)
        Chart chart = builder.InsertChart(ChartType.Column, 400, 300).Chart;
        // Populate chart with sample data
        chart.Series.Clear();
        ChartSeries series = chart.Series.Add("Series 1", new string[] { "A", "B", "C" }, new double[] { 10, 20, 30 });
        chart.Title.Text = "Sample Chart";

        // Save the document
        doc.Save(docPath);

        // Reload the document for extraction
        Document loadedDoc = new Document(docPath);

        // Iterate through all Shape nodes and extract images
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string outPath = Path.Combine(outputFolder, $"extracted-{extractedCount}.png");
                shape.ImageData.Save(outPath);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Optional cleanup (commented out)
        // File.Delete(imagePath);
        // File.Delete(docPath);
    }
}
