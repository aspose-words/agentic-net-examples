using System;
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

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Set a light gray background.
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Configure left border.
        Border leftBorder = customStyle.ParagraphFormat.Borders[BorderType.Left];
        leftBorder.LineStyle = LineStyle.Single;
        leftBorder.Color = Color.Black;
        leftBorder.LineWidth = 1.0;

        // Configure right border.
        Border rightBorder = customStyle.ParagraphFormat.Borders[BorderType.Right];
        rightBorder.LineStyle = LineStyle.Single;
        rightBorder.Color = Color.Black;
        rightBorder.LineWidth = 1.0;

        // Apply the custom style to a paragraph.
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("This paragraph uses a custom style with left/right borders and a light gray background.");

        // Save the document.
        string outputPath = "CustomStyle.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
