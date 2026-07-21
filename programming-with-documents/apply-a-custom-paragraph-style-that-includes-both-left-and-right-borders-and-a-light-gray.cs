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

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure left and right borders.
        Border leftBorder = builder.ParagraphFormat.Borders.Left;
        leftBorder.LineStyle = LineStyle.Single;
        leftBorder.LineWidth = 1.0;
        leftBorder.Color = Color.Black;

        Border rightBorder = builder.ParagraphFormat.Borders.Right;
        rightBorder.LineStyle = LineStyle.Single;
        rightBorder.LineWidth = 1.0;
        rightBorder.Color = Color.Black;

        // Set a light gray background shading.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Write the paragraph with the applied formatting.
        builder.Writeln("This paragraph has left/right borders and a light gray background.");

        // Save the document to a file in the current directory.
        doc.Save("CustomParagraphStyle.docx");
    }
}
