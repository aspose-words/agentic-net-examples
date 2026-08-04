using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // HTML string containing a table with cell background color and bold text.
        string html = @"
<table border='1' style='border-collapse:collapse;'>
    <tr>
        <td style='background-color:#FFCC00;'><b>Cell 1</b></td>
        <td>Cell 2</td>
    </tr>
</table>";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the HTML table into the document.
        builder.InsertHtml(html);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Reload the document to verify the table and its formatting.
        Document loadedDoc = new Document(outputPath);
        Table table = loadedDoc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
            throw new InvalidOperationException("No table was found in the document.");

        // Verify the first cell's background shading.
        Cell firstCell = table.Rows[0].Cells[0];
        Color expectedColor = Color.FromArgb(255, 255, 204, 0); // #FFCC00
        if (firstCell.CellFormat.Shading.BackgroundPatternColor.ToArgb() != expectedColor.ToArgb())
            throw new InvalidOperationException("Cell background color was not preserved.");

        // Verify the first cell contains bold text.
        Paragraph para = firstCell.FirstParagraph;
        Run run = para?.FirstChild as Run;
        if (run == null || !run.Font.Bold)
            throw new InvalidOperationException("Bold formatting of the cell text was not preserved.");

        Console.WriteLine("Table inserted from HTML and cell formatting preserved successfully.");
    }
}
