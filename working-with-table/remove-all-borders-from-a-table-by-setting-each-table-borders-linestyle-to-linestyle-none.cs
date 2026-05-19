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
        builder.EndTable();

        // Remove all borders by setting each border's line style to None.
        // Using SetBorders applies the style to every outer and inner border of the table.
        table.SetBorders(LineStyle.None, 0.0, Color.Empty);

        // Save the document.
        const string outputPath = "TableNoBorders.docx";
        doc.Save(outputPath);

        // Verify that the borders were removed.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
        Border topBorder = loadedTable.FirstRow.RowFormat.Borders[BorderType.Top];
        if (topBorder.LineStyle != LineStyle.None)
        {
            throw new InvalidOperationException("Failed to remove table borders.");
        }

        // Program ends without user interaction.
    }
}
