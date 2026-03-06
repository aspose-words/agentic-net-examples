using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline textbox shape with the desired size.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        textBox.WrapType = WrapType.Inline; // Ensure the shape is inline with text.

        // Add a centered paragraph inside the textbox.
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Add a run of text to the paragraph.
        Run run = new Run(doc);
        run.Text = "Hello Aspose!";
        para.AppendChild(run);

        // Append the shape to the document body.
        builder.CurrentParagraph.AppendChild(textBox);

        // Save the document to a DOCX file.
        doc.Save("ShapeWithText.docx");
    }
}
