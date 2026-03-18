using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class InsertTextBoxInFirstPageHeader
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Enable a different header for the first page.
        doc.FirstSection.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the builder to the first‑page header.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Create a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;          // No text wrapping – floating.
        textBox.Height = 50;                       // Height in points.
        textBox.Width = 200;                       // Width in points.
        textBox.HorizontalAlignment = HorizontalAlignment.Center;
        textBox.VerticalAlignment = VerticalAlignment.Top;

        // Add a paragraph inside the text box and put some text.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc);
        run.Text = "First page header text box";
        para.AppendChild(run);

        // Insert the text box into the header.
        builder.CurrentParagraph.AppendChild(textBox);

        // Save the document.
        doc.Save("HeaderFirstPageTextBox.docx");
    }
}
