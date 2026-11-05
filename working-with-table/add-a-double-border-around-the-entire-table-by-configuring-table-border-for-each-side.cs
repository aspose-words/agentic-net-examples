using System;
using System.IO;
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

        // Build a simple 3x3 table.
        builder.StartTable();

        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply a double line border to each side of the table.
        // SetBorder(BorderType, LineStyle, lineWidth, Color, true) adds the border.
        table.SetBorder(BorderType.Top, LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Left, LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Double, 1.5, Color.Black, true);

        // Save the document.
        string outputPath = "TableWithDoubleBorder.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
