using System;
using System.Drawing;
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

        // First row with two cells.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row with two cells.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a 2‑point single black border to all four sides of the table.
        // Table.Borders does not exist in this version of Aspose.Words;
        // use the SetBorders method instead.
        table.SetBorders(LineStyle.Single, 2.0, Color.Black);

        // Save the document to a file.
        const string outputPath = "TableBorders.docx";
        doc.Save(outputPath);
    }
}
