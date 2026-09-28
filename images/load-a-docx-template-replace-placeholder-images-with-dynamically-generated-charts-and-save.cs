using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class ImageChartReplacementExample
{
    public static void Main()
    {
        // Step 1: Create a sample template document with a placeholder image.
        const string placeholderImagePath = "placeholder.png";
        const string templatePath = "template.docx";
        const string placeholderAltText = "ChartPlaceholder";

        // Create a deterministic placeholder image.
        CreatePlaceholderImage(placeholderImagePath, 100, 100);

        // Build the template document and insert the placeholder image.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        Shape placeholderShape = builder.InsertImage(placeholderImagePath);
        placeholderShape.AlternativeText = placeholderAltText; // Mark for later replacement.
        templateDoc.Save(templatePath);

        // Step 2: Load the template document.
        Document loadedDoc = new Document(templatePath);

        // Step 3: Generate a chart image dynamically.
        const string chartImagePath = "chart.png";
        CreateSampleChartImage(chartImagePath, 500, 300);

        // Step 4: Replace placeholder images with the generated chart.
        bool replacementMade = false;
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.AlternativeText == placeholderAltText)
            {
                shape.ImageData.SetImage(chartImagePath);
                replacementMade = true;
            }
        }

        if (!replacementMade)
        {
            throw new InvalidOperationException("No placeholder image found to replace.");
        }

        // Step 5: Save the final document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Validation: ensure the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }

        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }

    // Creates a simple white placeholder PNG image.
    private static void CreatePlaceholderImage(string filePath, int width, int height)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightGray);
                // Optionally draw a cross to indicate placeholder.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkGray, 2))
                {
                    graphics.DrawLine(pen, 0, 0, width, height);
                    graphics.DrawLine(pen, width, 0, 0, height);
                }
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    // Generates a simple bar chart image.
    private static void CreateSampleChartImage(string filePath, int width, int height)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);

                // Define chart area.
                int chartLeft = 50;
                int chartBottom = height - 50;
                int chartTop = 50;
                int chartWidth = width - 100;
                int chartHeight = chartBottom - chartTop;

                // Draw axes.
                using (Aspose.Drawing.Pen axisPen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2))
                {
                    // Y axis.
                    graphics.DrawLine(axisPen, chartLeft, chartTop, chartLeft, chartBottom);
                    // X axis.
                    graphics.DrawLine(axisPen, chartLeft, chartBottom, chartLeft + chartWidth, chartBottom);
                }

                // Sample data for bars.
                int[] values = { 70, 40, 90, 55 };
                int barCount = values.Length;
                int maxValue = 100; // Assuming max value for scaling.

                int barSpacing = chartWidth / (barCount * 2);
                int barWidth = barSpacing;

                using (Aspose.Drawing.SolidBrush barBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.SteelBlue))
                {
                    for (int i = 0; i < barCount; i++)
                    {
                        int barHeight = (int)((values[i] / (float)maxValue) * chartHeight);
                        int x = chartLeft + barSpacing + i * (barWidth + barSpacing);
                        int y = chartBottom - barHeight;
                        graphics.FillRectangle(barBrush, x, y, barWidth, barHeight);
                    }
                }

                // Optional: draw value labels.
                using (Aspose.Drawing.SolidBrush textBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
                {
                    for (int i = 0; i < barCount; i++)
                    {
                        string label = values[i].ToString();
                        SizeF textSize = graphics.MeasureString(label, font);
                        int x = chartLeft + barSpacing + i * (barWidth + barSpacing) + (barWidth - (int)textSize.Width) / 2;
                        int y = chartBottom - (int)((values[i] / (float)maxValue) * chartHeight) - (int)textSize.Height - 5;
                        graphics.DrawString(label, font, textBrush, x, y);
                    }
                }
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }
}
