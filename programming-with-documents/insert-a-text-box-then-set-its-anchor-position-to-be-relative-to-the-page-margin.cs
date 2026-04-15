using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Set the anchor position to be relative to the page margins.
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;

        // Add some text inside the text box.
        builder.MoveTo(textBox.LastParagraph);
        builder.Write("Hello from the text box!");

        // Save the document.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TextBoxAnchor.docx");
        doc.Save(outputPath);
    }
}
