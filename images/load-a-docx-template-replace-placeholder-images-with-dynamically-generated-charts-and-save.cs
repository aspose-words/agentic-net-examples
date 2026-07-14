using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing provides Bitmap, Graphics, Color, SolidBrush, Font, Pen

public class Program
{
    // File names used in the example
    private const string TemplateFileName = "template.docx";
    private const string PlaceholderImageFile = "placeholder.png";
    private const string ChartImageFile = "chart.png";
    private const string ResultFileName = "result.docx";

    public static void Main()
    {
        // Ensure a clean start
        CleanupFiles();

        // 1. Create a placeholder image
        CreatePlaceholderImage();

        // 2. Build a DOCX template that contains the placeholder image
        CreateTemplateDocument();

        // 3. Load the template document
        Document doc = new Document(TemplateFileName);

        // 4. Generate a chart image dynamically
        CreateChartImage();

        // 5. Find the placeholder shape and replace its image with the chart
        ReplacePlaceholderWithChart(doc);

        // 6. Save the resulting document
        doc.Save(ResultFileName);

        // 7. Validate that the output file was created
        if (!File.Exists(ResultFileName))
            throw new InvalidOperationException($"Failed to create output file '{ResultFileName}'.");

        Console.WriteLine("Document created: " + Path.GetFullPath(ResultFileName));
    }

    private static void CleanupFiles()
    {
        foreach (var file in new[] { PlaceholderImageFile, TemplateFileName, ChartImageFile, ResultFileName })
        {
            if (File.Exists(file))
                File.Delete(file);
        }
    }

    private static void CreatePlaceholderImage()
    {
        // Create a 100x100 white bitmap with the word "PH" (placeholder) drawn on it
        int width = 100;
        int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24, FontStyle.Bold))
            {
                string text = "PH";
                var textSize = g.MeasureString(text, font);
                float x = (width - textSize.Width) / 2;
                float y = (height - textSize.Height) / 2;
                g.DrawString(text, font, new SolidBrush(Color.Gray), x, y);
            }
            bitmap.Save(PlaceholderImageFile);
        }
    }

    private static void CreateTemplateDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the placeholder image and mark it with AlternativeText so we can locate it later
        Shape placeholderShape = builder.InsertImage(PlaceholderImageFile);
        placeholderShape.AlternativeText = "PLACEHOLDER";

        // Add some surrounding text for clarity
        builder.Writeln();
        builder.Writeln("The image above will be replaced with a generated chart.");

        doc.Save(TemplateFileName);
    }

    private static void CreateChartImage()
    {
        // Simple bar chart: 3 bars with values 30, 70, 50
        int width = 400;
        int height = 300;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);

            // Draw axes
            using (Pen axisPen = new Pen(Color.Black, 2))
            {
                g.DrawLine(axisPen, 50, height - 50, width - 30, height - 50); // X axis
                g.DrawLine(axisPen, 50, height - 50, 50, 30); // Y axis
            }

            // Bar data
            int[] values = { 30, 70, 50 };
            Color[] barColors = { Color.Red, Color.Green, Color.Blue };
            int barWidth = 60;
            int spacing = 30;
            int maxBarHeight = height - 100; // leave margins

            for (int i = 0; i < values.Length; i++)
            {
                int barHeight = (int)((values[i] / 100.0) * maxBarHeight);
                int x = 70 + i * (barWidth + spacing);
                int y = height - 50 - barHeight;

                using (SolidBrush brush = new SolidBrush(barColors[i]))
                {
                    g.FillRectangle(brush, x, y, barWidth, barHeight);
                }

                // Draw value label above each bar
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
                {
                    string valText = values[i].ToString();
                    var textSize = g.MeasureString(valText, font);
                    float tx = x + (barWidth - textSize.Width) / 2;
                    float ty = y - textSize.Height - 5;
                    g.DrawString(valText, font, new SolidBrush(Color.Black), tx, ty);
                }
            }

            // X‑axis labels
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
            {
                string[] labels = { "A", "B", "C" };
                for (int i = 0; i < labels.Length; i++)
                {
                    var textSize = g.MeasureString(labels[i], font);
                    float tx = 70 + i * (barWidth + spacing) + (barWidth - textSize.Width) / 2;
                    float ty = height - 45;
                    g.DrawString(labels[i], font, new SolidBrush(Color.Black), tx, ty);
                }
            }

            bitmap.Save(ChartImageFile);
        }
    }

    private static void ReplacePlaceholderWithChart(Document doc)
    {
        // Locate the shape that has the placeholder AlternativeText
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.AlternativeText == "PLACEHOLDER")
            {
                // Replace the image data with the generated chart image
                shape.ImageData.SetImage(ChartImageFile);
                break;
            }
        }
    }
}
