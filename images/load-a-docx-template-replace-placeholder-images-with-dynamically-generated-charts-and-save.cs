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
        // File names
        const string placeholderImagePath = "placeholder.png";
        const string templatePath = "template.docx";
        const string chartImagePath = "chart.png";
        const string outputPath = "output.docx";

        // 1. Create a placeholder image (white background with text)
        CreatePlaceholderImage(placeholderImagePath);

        // 2. Build a DOCX template containing the placeholder image
        CreateTemplateDocument(templatePath, placeholderImagePath);

        // 3. Load the template document
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 4. Generate a chart image
        CreateChartImage(chartImagePath);

        // 5. Replace each placeholder image with the generated chart
        bool replaced = false;
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.AlternativeText == "ChartPlaceholder")
            {
                shape.ImageData.SetImage(chartImagePath);
                replaced = true;
            }
        }

        if (!replaced)
            throw new InvalidOperationException("No placeholder image found to replace.");

        // 6. Save the resulting document
        doc.Save(outputPath);

        // 7. Validate output
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    private static void CreatePlaceholderImage(string path)
    {
        const int width = 200;
        const int height = 150;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2))
                {
                    g.DrawRectangle(pen, 0, 0, width - 1, height - 1);
                }
                // Use fully qualified Aspose.Drawing.Font to avoid ambiguity
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
                {
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                    {
                        g.DrawString("Placeholder", font, brush, new Aspose.Drawing.PointF(20, height / 2 - 10));
                    }
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    private static void CreateTemplateDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert placeholder image
        Shape placeholderShape = builder.InsertImage(imagePath);
        placeholderShape.AlternativeText = "ChartPlaceholder";

        // Add some text after the image
        builder.Writeln();
        builder.Writeln("This document contains a placeholder image that will be replaced by a chart.");

        doc.Save(docPath);
    }

    private static void CreateChartImage(string path)
    {
        const int width = 400;
        const int height = 300;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);

                // Draw axes
                using (Aspose.Drawing.Pen axisPen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2))
                {
                    // Y axis
                    g.DrawLine(axisPen, 50, 20, 50, height - 50);
                    // X axis
                    g.DrawLine(axisPen, 50, height - 50, width - 20, height - 50);
                }

                // Sample data
                int[] values = { 70, 120, 90, 150, 110 };
                Aspose.Drawing.Color[] barColors = {
                    Aspose.Drawing.Color.Red,
                    Aspose.Drawing.Color.Green,
                    Aspose.Drawing.Color.Blue,
                    Aspose.Drawing.Color.Orange,
                    Aspose.Drawing.Color.Purple
                };

                int barWidth = 40;
                int spacing = 30;
                int maxBarHeight = height - 100;
                int maxValue = 160; // for scaling

                for (int i = 0; i < values.Length; i++)
                {
                    int barHeight = (int)((values[i] / (float)maxValue) * maxBarHeight);
                    int x = 50 + spacing + i * (barWidth + spacing);
                    int y = height - 50 - barHeight;

                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(barColors[i]))
                    {
                        g.FillRectangle(brush, x, y, barWidth, barHeight);
                    }

                    using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 1))
                    {
                        g.DrawRectangle(pen, x, y, barWidth, barHeight);
                    }
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}
