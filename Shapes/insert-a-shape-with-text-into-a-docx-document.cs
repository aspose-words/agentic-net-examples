using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ShapeWithTextExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert an inline TextBox shape with the desired size (width, height in points).
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

            // Set the shape to have no text wrapping (optional, can be adjusted as needed).
            textBox.WrapType = WrapType.None;

            // Access the first paragraph inside the shape.
            Paragraph shapeParagraph = textBox.FirstParagraph;

            // Create a run with the text you want to display inside the shape.
            Run run = new Run(doc);
            run.Text = "Hello Aspose!";

            // Append the run to the shape's paragraph.
            shapeParagraph.AppendChild(run);

            // Save the document to a DOCX file.
            doc.Save("ShapeWithText.docx");
        }
    }
}
