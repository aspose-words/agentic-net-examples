using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string artifactsDir = "Artifacts";
        const string templatePath = "Artifacts/template.docx";
        const string placeholderImagePath = "Artifacts/placeholder.png";
        const string chartImagePath = "Artifacts/chart.png";
        const string outputPath = "Artifacts/output.docx";

        // Ensure output directory exists
        Directory.CreateDirectory(artifactsDir);

        // -------------------------------------------------
        // 1. Create a placeholder image (used in the template)
        // -------------------------------------------------
        const int placeholderWidth = 100;
        const int placeholderHeight = 100;
        using (Aspose.Drawing.Bitmap placeholderBitmap = new Aspose.Drawing.Bitmap(placeholderWidth, placeholderHeight))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(placeholderBitmap))
        {
            g.Clear(Aspose.Drawing.Color.White);
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
            {
                g.DrawString("PH", font, new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Gray), new Aspose.Drawing.PointF(20, 40));
            }
            placeholderBitmap.Save(placeholderImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Build a DOCX template containing the placeholder image
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert the placeholder image and mark it with AlternativeText for later identification
        Shape placeholderShape = builder.InsertImage(placeholderImagePath);
        placeholderShape.AlternativeText = "ChartPlaceholder";

        // Save the template
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 3. Generate a simple chart image dynamically
        // -------------------------------------------------
        const int chartWidth = 400;
        const int chartHeight = 300;
        using (Aspose.Drawing.Bitmap chartBitmap = new Aspose.Drawing.Bitmap(chartWidth, chartHeight))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(chartBitmap))
        {
            g.Clear(Aspose.Drawing.Color.White);

            // Draw axes
            Aspose.Drawing.Pen axisPen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2);
            g.DrawLine(axisPen, 50, chartHeight - 50, chartWidth - 30, chartHeight - 50); // X axis
            g.DrawLine(axisPen, 50, chartHeight - 50, 50, 30); // Y axis

            // Sample data
            int[] values = { 70, 120, 90, 150 };
            Aspose.Drawing.Color[] barColors = { Aspose.Drawing.Color.Blue, Aspose.Drawing.Color.Green, Aspose.Drawing.Color.Orange, Aspose.Drawing.Color.Purple };
            int barWidth = 40;
            int spacing = 20;
            int maxVal = 160; // for scaling

            // Draw bars
            for (int i = 0; i < values.Length; i++)
            {
                int barHeight = (int)((values[i] / (float)maxVal) * (chartHeight - 100));
                int x = 50 + spacing + i * (barWidth + spacing);
                int y = chartHeight - 50 - barHeight;
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(barColors[i]))
                {
                    g.FillRectangle(brush, x, y, barWidth, barHeight);
                }
                // Draw value label
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 10))
                {
                    g.DrawString(values[i].ToString(), font, new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black), new Aspose.Drawing.PointF(x, y - 15));
                }
            }

            // Draw X-axis labels
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 10))
            {
                string[] labels = { "Q1", "Q2", "Q3", "Q4" };
                for (int i = 0; i < labels.Length; i++)
                {
                    int x = 50 + spacing + i * (barWidth + spacing) + barWidth / 2 - 10;
                    g.DrawString(labels[i], font, new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black), new Aspose.Drawing.PointF(x, chartHeight - 45));
                }
            }

            chartBitmap.Save(chartImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -------------------------------------------------
        // 4. Load the template and replace the placeholder with the chart image
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        bool replaced = false;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.AlternativeText == "ChartPlaceholder")
            {
                // Replace image data with the generated chart
                shape.ImageData.SetImage(chartImagePath);
                replaced = true;
                break; // Assuming only one placeholder
            }
        }

        if (!replaced)
            throw new InvalidOperationException("Placeholder shape not found in the document.");

        // -------------------------------------------------
        // 5. Save the final document
        // -------------------------------------------------
        doc.Save(outputPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 6. Validate that the output file exists
        // -------------------------------------------------
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
