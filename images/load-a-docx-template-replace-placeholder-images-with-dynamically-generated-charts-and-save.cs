using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // File paths
        const string placeholderImagePath = "placeholder.png";
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -------------------------------------------------
        // 1. Create a deterministic placeholder image file
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 100;

        // Use Aspose.Drawing types to avoid System.Drawing ambiguity
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        graphics.DrawString(
            "Placeholder",
            new Aspose.Drawing.Font("Arial", 12),
            new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
            new Aspose.Drawing.PointF(10, 40));
        graphics.Dispose();

        // Save the placeholder image to disk
        bitmap.Save(placeholderImagePath);
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Build a template document containing the placeholder image
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.InsertImage(placeholderImagePath);
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 3. Load the template document
        // -------------------------------------------------
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 4. Locate the placeholder shape (first shape with an image)
        // -------------------------------------------------
        Shape placeholderShape = null;
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                placeholderShape = shape;
                break;
            }
        }

        if (placeholderShape == null)
            throw new Exception("Placeholder image not found in the template.");

        // -------------------------------------------------
        // 5. Replace the placeholder with a dynamically generated chart
        // -------------------------------------------------
        // Move the builder cursor to the placeholder shape
        builder.MoveTo(placeholderShape);

        // Insert a column chart with the desired size (points)
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Clear any default series and add our own data
        chart.Series.Clear();
        chart.Series.Add(
            "Sample Series",
            new[] { "Item 1", "Item 2", "Item 3" },
            new double[] { 10, 30, 20 });

        // Remove the original placeholder shape
        placeholderShape.Remove();

        // -------------------------------------------------
        // 6. Save the resulting document
        // -------------------------------------------------
        doc.Save(outputPath);

        // -------------------------------------------------
        // 7. Validate that the output file was created
        // -------------------------------------------------
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Cleanup temporary files (optional)
        File.Delete(placeholderImagePath);
        File.Delete(templatePath);
    }
}
