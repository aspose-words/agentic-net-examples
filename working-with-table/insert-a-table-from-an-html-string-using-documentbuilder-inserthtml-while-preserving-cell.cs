using System;
using System.IO;
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

        // HTML string containing a table with cell formatting (background color and bold text).
        string htmlTable = @"
            <table border='1' style='border-collapse:collapse;'>
                <tr>
                    <td style='background-color:#FFFF00; font-weight:bold;'>Header 1</td>
                    <td style='background-color:#00FF00; font-weight:bold;'>Header 2</td>
                </tr>
                <tr>
                    <td style='background-color:#FFCCCC;'>Cell 1</td>
                    <td style='background-color:#CCCCFF;'>Cell 2</td>
                </tr>
            </table>";

        // Insert the HTML table into the document.
        builder.InsertHtml(htmlTable);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved.");

        // Reload the document to verify table and cell formatting.
        Document loadedDoc = new Document(outputPath);
        Table table = loadedDoc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
            throw new Exception("No table was found in the saved document.");

        // Verify the first cell's background color (should be Yellow).
        Cell firstCell = table.Rows[0].Cells[0];
        Color expectedColor = Color.Yellow; // #FFFF00
        if (firstCell.CellFormat.Shading.BackgroundPatternColor.ToArgb() != expectedColor.ToArgb())
            throw new Exception("Cell formatting was not preserved for the first cell.");

        // Verify the second cell's background color (should be Lime).
        Cell secondCell = table.Rows[0].Cells[1];
        expectedColor = Color.Lime; // #00FF00
        if (secondCell.CellFormat.Shading.BackgroundPatternColor.ToArgb() != expectedColor.ToArgb())
            throw new Exception("Cell formatting was not preserved for the second cell.");

        // If execution reaches this point, the table was inserted and formatting preserved.
        Console.WriteLine("Table inserted from HTML and cell formatting preserved successfully.");
    }
}
