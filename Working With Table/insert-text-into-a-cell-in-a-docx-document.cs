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

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the cell we want to modify (second row, first column).
        Cell targetCell = table.Rows[1].Cells[0];

        // Move the builder cursor to the beginning of that cell.
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert additional text into the cell.
        builder.Write(" - Inserted text");

        // Save the document.
        doc.Save("Output.docx");
    }
}
