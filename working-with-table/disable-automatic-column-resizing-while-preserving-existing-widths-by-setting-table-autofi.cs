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

        // Insert first cell and set a fixed preferred width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Fixed width cell 1");

        // Insert second cell and set a fixed preferred width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Fixed width cell 2");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Disable automatic column resizing while preserving the existing column widths.
        // This can be done by turning off the AllowAutoFit flag.
        table.AllowAutoFit = false;

        // Save the document to a file.
        const string outputFile = "TableAutoFitDisabled.docx";
        doc.Save(outputFile);
    }
}
