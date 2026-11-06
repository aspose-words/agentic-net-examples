using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x3 table.
        Table table = builder.StartTable();
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

        // Insert a text watermark that will appear behind the content of every cell,
        // including each cell of the first row.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the resulting document.
        string outputPath = "WatermarkTableCell.docx";
        doc.Save(outputPath);
    }
}
