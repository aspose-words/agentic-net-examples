using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // First cell: set vertical alignment via CellFormat and horizontal alignment via ParagraphFormat.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Centered");

        // Second cell: apply the same alignment.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Centered");

        // End the first row.
        builder.EndRow();

        // Validate that the alignments were applied.
        if (table.Rows[0].Cells[0].CellFormat.VerticalAlignment != CellVerticalAlignment.Center ||
            table.Rows[0].Cells[0].FirstParagraph.ParagraphFormat.Alignment != ParagraphAlignment.Center ||
            table.Rows[0].Cells[1].CellFormat.VerticalAlignment != CellVerticalAlignment.Center ||
            table.Rows[0].Cells[1].FirstParagraph.ParagraphFormat.Alignment != ParagraphAlignment.Center)
        {
            throw new InvalidOperationException("Cell alignment was not set correctly.");
        }

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save("AlignedTable.docx");
    }
}
