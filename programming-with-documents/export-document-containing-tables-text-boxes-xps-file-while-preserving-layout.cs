using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Insert a 2x2 table ----------
        builder.StartTable();

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

        // ---------- Insert a text box ----------
        // Create a text box shape of width 200 points and height 100 points.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Move the cursor inside the text box and add some text.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("Text inside the box");

        // ---------- Save as XPS ----------
        // Create XpsSaveOptions. The default OptimizeOutput = false preserves the original layout.
        XpsSaveOptions saveOptions = new XpsSaveOptions
        {
            OptimizeOutput = false // explicit for clarity
        };

        // Save the document to an XPS file.
        doc.Save("DocumentWithTableAndTextBox.xps", saveOptions);
    }
}
