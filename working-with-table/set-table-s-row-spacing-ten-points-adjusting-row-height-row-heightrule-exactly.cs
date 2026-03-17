using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // ----- First row -----
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Set the first row's height to exactly 10 points.
        table.Rows[0].RowFormat.Height = 10;
        table.Rows[0].RowFormat.HeightRule = HeightRule.Exactly;

        // ----- Second row -----
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Set the second row's height to exactly 10 points.
        table.Rows[1].RowFormat.Height = 10;
        table.Rows[1].RowFormat.HeightRule = HeightRule.Exactly;

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save("TableRowSpacing.docx");
    }
}
