using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsShapeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define textbox dimensions (points).
            double width = 200;   // 200 points ≈ 2.78 inches
            double height = 100;  // 100 points ≈ 1.39 inches

            // Insert a floating textbox shape.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, width, height);
            textBox.WrapType = WrapType.None;
            textBox.HorizontalAlignment = HorizontalAlignment.Center;
            textBox.VerticalAlignment = VerticalAlignment.Top;

            // Apply border (stroke) formatting.
            textBox.Stroke.Color = Color.DarkBlue;      // Border color
            textBox.StrokeWeight = 2.0;                // Border thickness (points)
            textBox.Stroke.DashStyle = DashStyle.Dash; // Dashed line style

            // Apply interior fill formatting.
            textBox.FillColor = Color.LightYellow;

            // Add a paragraph with centered text inside the textbox.
            textBox.AppendChild(new Paragraph(doc));
            Paragraph para = textBox.FirstParagraph;
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            Run run = new Run(doc) { Text = "Hello Aspose.Words!" };
            para.AppendChild(run);

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TextboxShape.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Document was not saved correctly.");
        }
    }
}
