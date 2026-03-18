using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Set vertical alignment of text inside the cells to middle.
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

        // Set horizontal alignment of paragraphs (cell contents) to center.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // First cell.
        builder.InsertCell();
        builder.Write("Centered Text 1");

        // Second cell.
        builder.InsertCell();
        builder.Write("Centered Text 2");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("CellAlignmentCentered.docx");
    }
}
