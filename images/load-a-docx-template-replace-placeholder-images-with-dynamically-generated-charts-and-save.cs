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
        // File names
        const string placeholderImagePath = "placeholder.png";
        const string chartImagePath = "chart.png";
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -------------------------------------------------
        // Step 1: Create a placeholder image (simple X)
        // -------------------------------------------------
        using (Bitmap placeholderBitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(placeholderBitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Black, 5))
                {
                    g.DrawLine(pen, 0, 0, 99, 99);
                    g.DrawLine(pen, 0, 99, 99, 0);
                }
            }
            placeholderBitmap.Save(placeholderImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 2: Create a DOCX template and insert placeholder image
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // Insert placeholder image
        builder.InsertImage(placeholderImagePath);
        // Retrieve the inserted shape and assign a name for later identification
        Shape placeholderShape = builder.CurrentParagraph.LastChild as Shape;
        if (placeholderShape == null || !placeholderShape.HasImage)
            throw new InvalidOperationException("Failed to insert placeholder image.");
        placeholderShape.Name = "ChartPlaceholder";
        // Save the template document
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Step 3: Load the template document
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder docBuilder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Step 4: Dynamically generate a chart image
        // -------------------------------------------------
        using (Bitmap chartBitmap = new Bitmap(400, 300))
        {
            using (Graphics g = Graphics.FromImage(chartBitmap))
            {
                g.Clear(Color.White);
                // Draw axes
                using (Pen axisPen = new Pen(Color.Black, 2))
                {
                    g.DrawLine(axisPen, 50, 250, 350, 250); // X axis
                    g.DrawLine(axisPen, 50, 250, 50, 50);   // Y axis
                }
                // Sample data
                int[] values = { 70, 120, 90, 150 };
                Color[] barColors = { Color.Red, Color.Green, Color.Blue, Color.Orange };
                int barWidth = 40;
                int spacing = 20;
                for (int i = 0; i < values.Length; i++)
                {
                    int barHeight = (int)((values[i] / 150.0) * 180); // Scale to max height 180
                    int x = 60 + i * (barWidth + spacing);
                    int y = 250 - barHeight;
                    using (Brush brush = new SolidBrush(barColors[i]))
                    {
                        g.FillRectangle(brush, x, y, barWidth, barHeight);
                    }
                    using (Pen pen = new Pen(Color.Black, 1))
                    {
                        g.DrawRectangle(pen, x, y, barWidth, barHeight);
                    }
                }
            }
            chartBitmap.Save(chartImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 5: Replace placeholder image with generated chart
        // -------------------------------------------------
        bool replaced = false;
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.Name == "ChartPlaceholder")
            {
                shape.ImageData.SetImage(chartImagePath);
                replaced = true;
                break;
            }
        }
        if (!replaced)
            throw new InvalidOperationException("Placeholder shape not found in the document.");

        // -------------------------------------------------
        // Step 6: Save the final document
        // -------------------------------------------------
        doc.Save(outputPath);

        // -------------------------------------------------
        // Validation: Ensure output file exists
        // -------------------------------------------------
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
