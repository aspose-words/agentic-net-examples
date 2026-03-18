using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape. Width = 200 points, Height = 100 points.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Optional: make the text box float without wrapping surrounding text.
        textBox.WrapType = WrapType.None;

        // Move the builder's cursor to the last paragraph inside the text box.
        builder.MoveTo(textBox.LastParagraph);

        // Start building a table inside the text box.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save("FloatingTextboxWithTable.docx");
    }
}
