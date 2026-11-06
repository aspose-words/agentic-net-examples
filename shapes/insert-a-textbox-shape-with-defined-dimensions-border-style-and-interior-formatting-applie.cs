using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox shape with defined dimensions.
        double width = 200;   // Width in points.
        double height = 100;  // Height in points.
        Shape textbox = builder.InsertShape(ShapeType.TextBox, width, height);

        // Apply border (stroke) formatting.
        textbox.Stroke.Color = Color.Black;
        textbox.Stroke.Weight = 1.5;               // Border thickness in points.
        textbox.Stroke.DashStyle = DashStyle.Dash; // Dashed border.

        // Apply interior fill formatting.
        textbox.Fill.Color = Color.LightYellow;

        // Insert text into the textbox.
        builder.MoveTo(textbox.FirstParagraph);
        builder.Font.Size = 12;
        builder.Write("This is a textbox shape.");

        // Save the document to disk.
        string outputPath = "TextboxShape.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
