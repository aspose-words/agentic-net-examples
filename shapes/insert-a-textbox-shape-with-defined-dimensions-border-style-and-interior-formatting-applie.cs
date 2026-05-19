using System;
using System.IO;
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

            // Insert a floating text box shape with specific dimensions (200x100 points).
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

            // Configure the shape's border (stroke) style.
            textBox.StrokeColor = System.Drawing.Color.DarkBlue;          // Border color.
            textBox.StrokeWeight = 2.0;                                   // Border thickness (points).
            textBox.Stroke.DashStyle = DashStyle.Dash;                    // Dashed line.

            // Configure the interior fill of the shape.
            textBox.FillColor = System.Drawing.Color.LightYellow;

            // Position the shape as a floating object.
            textBox.WrapType = WrapType.None;
            textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBox.Left = 100;   // Distance from the left edge of the page (points).
            textBox.Top = 150;    // Distance from the top edge of the page (points).

            // Add a paragraph inside the text box and some sample text.
            textBox.AppendChild(new Paragraph(doc));
            Paragraph innerParagraph = textBox.FirstParagraph;
            innerParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            Run run = new Run(doc);
            run.Text = "Hello Aspose.Words!";
            innerParagraph.AppendChild(run);

            // Save the document to the local file system.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "TextBoxShape.docx");
            doc.Save(outputPath);

            // Validate that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");

            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
