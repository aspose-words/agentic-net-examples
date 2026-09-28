using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // HTML string containing a table with cell background shading.
        string html = @"
<table border='1' style='border-collapse:collapse;'>
    <tr>
        <td style='background-color:#FFFF00;'>Yellow Cell</td>
        <td>Normal Cell</td>
    </tr>
    <tr>
        <td>Cell 3</td>
        <td>Cell 4</td>
    </tr>
</table>";

        // Insert the HTML into the document.
        builder.Writeln("Table inserted from HTML:");
        builder.InsertHtml(html);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Load the saved document to verify the table and cell formatting.
        Document loadedDoc = new Document(outputPath);
        Table table = loadedDoc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
            throw new Exception("Table was not inserted.");

        Cell firstCell = table.Rows[0].Cells[0];
        Color bgColor = firstCell.CellFormat.Shading.BackgroundPatternColor;
        if (bgColor.ToArgb() != Color.Yellow.ToArgb())
            throw new Exception("Cell shading was not preserved.");

        // Indicate success (optional).
        Console.WriteLine("Table inserted and cell formatting preserved successfully.");
    }
}
