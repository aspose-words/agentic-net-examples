using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Add a single cell with some text.
        builder.InsertCell();
        builder.Writeln("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Apply a custom color to the top border of the table.
        // Use SetBorder because Table does not expose a Borders collection directly.
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.0, Color.Blue, true);

        // Save the document.
        doc.Save("TableTopBorderColor.docx");
    }
}
