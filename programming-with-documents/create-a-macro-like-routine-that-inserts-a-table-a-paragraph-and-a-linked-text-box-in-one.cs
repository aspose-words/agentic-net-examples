using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Perform the macro‑like operation: insert a table, a paragraph, and a linked text box.
        InsertTableParagraphAndLinkedTextBox(doc);

        // Save the resulting document.
        string outputPath = "MacroLikeOutput.docx";
        doc.Save(outputPath);
    }

    private static void InsertTableParagraphAndLinkedTextBox(Document doc)
    {
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Insert a 2x2 table -----
        Table table = builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // ----- Insert a paragraph after the table -----
        builder.Writeln(); // Ensure a line break after the table.
        builder.Writeln("This is a paragraph inserted after the table.");

        // ----- Insert a linked text box -----
        // Create a shape of type TextBox.
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 200,
            Height = 100,
            WrapType = WrapType.Inline
        };

        // Add text inside the text box.
        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "Linked Text Box Content");
        tbParagraph.AppendChild(tbRun);
        textBox.AppendChild(tbParagraph);

        // Insert the text box into the document.
        builder.InsertNode(textBox);
    }
}
