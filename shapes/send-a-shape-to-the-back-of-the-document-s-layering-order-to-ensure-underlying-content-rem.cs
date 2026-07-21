using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first rectangle (will be sent to the back).
        Shape backShape = builder.InsertShape(ShapeType.Rectangle, 200, 200);
        backShape.FillColor = System.Drawing.Color.LightBlue;
        backShape.Left = 100;   // Position to overlap with the second shape.
        backShape.Top = 100;

        // Insert the second rectangle (will stay in front).
        Shape frontShape = builder.InsertShape(ShapeType.Rectangle, 200, 200);
        frontShape.FillColor = System.Drawing.Color.OrangeRed;
        frontShape.Left = 150;  // Overlap the first shape.
        frontShape.Top = 150;

        // Send the first shape to the back by setting a lower ZOrder.
        backShape.ZOrder = 0;   // Lowest stacking order.
        frontShape.ZOrder = 1;  // Higher stacking order, appears in front.

        // Define output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ShapeSendToBack.docx");

        // Save the document.
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
