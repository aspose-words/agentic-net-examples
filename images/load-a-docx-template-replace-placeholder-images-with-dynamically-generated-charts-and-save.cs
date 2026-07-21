using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // File names used in the example.
        const string placeholderImagePath = "placeholder.png";
        const string chartImagePath = "chart.png";
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -------------------------------------------------
        // 1. Create a placeholder image (simple white box).
        // -------------------------------------------------
        const int placeholderWidth = 200;
        const int placeholderHeight = 150;
        using (Aspose.Drawing.Bitmap placeholderBitmap = new Aspose.Drawing.Bitmap(placeholderWidth, placeholderHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(placeholderBitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Gray, 2))
                {
                    g.DrawRectangle(pen, 0, 0, placeholderWidth - 1, placeholderHeight - 1);
                }
            }
            placeholderBitmap.Save(placeholderImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Build a DOCX template that contains the placeholder image.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // Insert the placeholder image and mark it with AlternativeText for later identification.
        Shape placeholderShape = builder.InsertImage(placeholderImagePath);
        placeholderShape.AlternativeText = "ChartPlaceholder";
        // Save the template.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 3. Dynamically generate a chart image (simple bar chart).
        // -------------------------------------------------
        const int chartWidth = 400;
        const int chartHeight = 300;
        using (Aspose.Drawing.Bitmap chartBitmap = new Aspose.Drawing.Bitmap(chartWidth, chartHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(chartBitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);

                // Sample data.
                int[] values = { 70, 45, 90, 55 };
                Aspose.Drawing.Color[] barColors = {
                    Aspose.Drawing.Color.Blue,
                    Aspose.Drawing.Color.Green,
                    Aspose.Drawing.Color.Orange,
                    Aspose.Drawing.Color.Purple
                };
                int barCount = values.Length;
                int maxValue = values.Max();

                int margin = 40;
                int availableWidth = chartWidth - 2 * margin;
                int barWidth = availableWidth / (barCount * 2);
                int spacing = barWidth; // equal spacing between bars.

                for (int i = 0; i < barCount; i++)
                {
                    int barHeight = (int)((values[i] / (float)maxValue) * (chartHeight - 2 * margin));
                    int x = margin + i * (barWidth + spacing);
                    int y = chartHeight - margin - barHeight;

                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(barColors[i]))
                    {
                        g.FillRectangle(brush, x, y, barWidth, barHeight);
                    }

                    // Draw value label.
                    using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 10))
                    using (Aspose.Drawing.SolidBrush textBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                    {
                        string valueStr = values[i].ToString();
                        SizeF textSize = g.MeasureString(valueStr, font);
                        float textX = x + (barWidth - textSize.Width) / 2;
                        float textY = y - textSize.Height - 2;
                        g.DrawString(valueStr, font, textBrush, textX, textY);
                    }
                }

                // Draw axes.
                using (Aspose.Drawing.Pen axisPen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2))
                {
                    g.DrawLine(axisPen, margin, margin, margin, chartHeight - margin); // Y axis
                    g.DrawLine(axisPen, margin, chartHeight - margin, chartWidth - margin, chartHeight - margin); // X axis
                }
            }
            chartBitmap.Save(chartImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -------------------------------------------------
        // 4. Load the template and replace the placeholder with the chart.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // Find all shapes that have the placeholder AlternativeText.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        bool replacementMade = false;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.AlternativeText == "ChartPlaceholder")
            {
                // Replace the image data with the generated chart.
                shape.ImageData.SetImage(chartImagePath);
                // Adjust the shape size to match the new image's original dimensions.
                shape.Width = shape.ImageData.ImageSize.WidthPoints;
                shape.Height = shape.ImageData.ImageSize.HeightPoints;
                replacementMade = true;
            }
        }

        if (!replacementMade)
            throw new InvalidOperationException("Placeholder image not found in the document.");

        // Save the final document.
        doc.Save(outputPath);

        // -------------------------------------------------
        // 5. Validate that the output file was created.
        // -------------------------------------------------
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
