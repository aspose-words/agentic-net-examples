using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a custom shape: Single corner snipped rectangle.
        // Width and height are specified in points (1 point = 1/72 inch).
        double shapeWidth = 200; // points
        double shapeHeight = 100; // points
        Shape shape = builder.InsertShape(ShapeType.SingleCornerSnipped, shapeWidth, shapeHeight);

        // Set a light gray fill color.
        shape.FillColor = Color.LightGray;

        // Define the corner radius via the Adjustments collection.
        // For snipped corner shapes the first adjustment controls the snip size.
        // The Adjustments collection is read‑only, so we modify the Adjustment object's Value.
        if (shape.Adjustments.Count > 0)
        {
            shape.Adjustments[0].Value = 10; // 10 points snip size
        }

        // Optionally, remove the outline stroke.
        shape.Stroke.On = false;

        // Save the document with a compliance level that supports DML shapes.
        string fileName = "SnipCornerRectangle.docx";
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save(filePath, saveOptions);

        // Validate that the file was created.
        if (!File.Exists(filePath))
            throw new Exception($"Failed to create the output file: {filePath}");
    }
}
