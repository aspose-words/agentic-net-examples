using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // The default document already contains a Section, Body, and a Paragraph.
        // Create a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;          // No text wrapping.
        textBox.Width = 200;                       // Width in points.
        textBox.Height = 50;                       // Height in points.
        textBox.HorizontalAlignment = HorizontalAlignment.Center;
        textBox.VerticalAlignment = VerticalAlignment.Top;

        // Add a paragraph inside the text box to hold text.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph innerParagraph = textBox.FirstParagraph;
        innerParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Add a run of text to the inner paragraph.
        Run run = new Run(doc);
        run.Text = "Hello world!";
        innerParagraph.AppendChild(run);

        // Insert the shape into the document's first paragraph.
        doc.FirstSection.Body.FirstParagraph.AppendChild(textBox);

        // Save the document to a .docx file.
        doc.Save("ShapeManipulation.docx");
    }
}
