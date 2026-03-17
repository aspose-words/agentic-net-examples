using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        builder.StartTable();

        // First cell – set text direction to vertical for East Asian characters.
        builder.InsertCell();
        // TextOrientation.VerticalFarEast makes Far East characters appear vertically.
        builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
        builder.Write("こんにちは"); // Japanese greeting displayed vertically.

        // Second cell – normal (horizontal) text direction.
        builder.InsertCell();
        builder.Write("Hello");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("CellVerticalEastAsian.docx");
    }
}
