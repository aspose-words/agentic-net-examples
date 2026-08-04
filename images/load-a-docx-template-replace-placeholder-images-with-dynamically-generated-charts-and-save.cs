using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders and file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string placeholderImagePath = Path.Combine(artifactsDir, "placeholder.png");
        string chartImagePath = Path.Combine(artifactsDir, "chart.png");
        string templatePath = Path.Combine(artifactsDir, "Template.docx");
        string resultPath = Path.Combine(artifactsDir, "Result.docx");

        // -------------------------------------------------
        // 1. Create a placeholder image (white square).
        // -------------------------------------------------
        CreateSampleImage(placeholderImagePath, 100, 100, Aspose.Drawing.Color.White);

        // -------------------------------------------------
        // 2. Build a template document that contains the placeholder image.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Document with a placeholder image that will be replaced by a chart:");
        Shape placeholderShape = builder.InsertImage(placeholderImagePath);
        placeholderShape.AlternativeText = "ChartPlaceholder"; // marker for replacement
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 3. Load the template document.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // 4. Dynamically generate a chart image.
        // -------------------------------------------------
        CreateSampleChart(chartImagePath, 400, 300);

        // -------------------------------------------------
        // 5. Replace each placeholder image with the generated chart.
        // -------------------------------------------------
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .OfType<Shape>()
                        .Where(s => s.HasImage && s.AlternativeText == "ChartPlaceholder");

        foreach (Shape shape in shapes)
        {
            shape.ImageData.SetImage(chartImagePath);
        }

        // -------------------------------------------------
        // 6. Save the final document.
        // -------------------------------------------------
        doc.Save(resultPath);

        // -------------------------------------------------
        // 7. Validate that the output file exists.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new Exception("Result document was not created.");

        Console.WriteLine("Document created successfully at: " + resultPath);
    }

    // Creates a simple bitmap image with a solid background color.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(backColor);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Generates a very simple bar‑chart‑like image.
    private static void CreateSampleChart(string filePath, int width, int height)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);

        // Define bar dimensions.
        int[] values = { 70, 40, 90, 55 };
        int barWidth = width / (values.Length * 2);
        int maxBarHeight = height - 40; // leave margin

        // Draw axes.
        using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2))
        {
            g.DrawLine(pen, 30, height - 30, width - 10, height - 30); // X axis
            g.DrawLine(pen, 30, height - 30, 30, 10); // Y axis
        }

        // Draw bars.
        for (int i = 0; i < values.Length; i++)
        {
            int barHeight = (int)((values[i] / 100.0) * maxBarHeight);
            int x = 30 + (i * 2 + 1) * barWidth;
            int y = height - 30 - barHeight;
            using (var brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Blue))
            {
                g.FillRectangle(brush, x, y, barWidth, barHeight);
            }
        }

        // Save the chart image.
        bitmap.Save(filePath, ImageFormat.Png);
        g.Dispose();
        bitmap.Dispose();
    }
}
