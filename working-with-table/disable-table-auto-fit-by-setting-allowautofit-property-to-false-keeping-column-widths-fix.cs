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

        // Start building a table.
        Table table = builder.StartTable();

        // First cell – set a fixed width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Fixed width cell 1");

        // Second cell – set a different fixed width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Fixed width cell 2");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Disable automatic table autofit to keep column widths fixed.
        table.AllowAutoFit = false;

        // Save the document to a file.
        doc.Save("TableAllowAutoFit.docx");
    }
}
