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

        // HTML string containing a table with cell background colors and borders.
        string html = @"
<table style='border-collapse:collapse;'>
    <tr>
        <td style='background-color:#FFFF00; border:2px solid #FF0000;'>Yellow cell</td>
        <td style='background-color:#00FF00; border:2px solid #0000FF;'>Green cell</td>
    </tr>
    <tr>
        <td style='background-color:#CCCCCC; border:2px solid #000000;'>Gray cell</td>
        <td style='border:2px solid #000000;'>No fill</td>
    </tr>
</table>";

        // Insert the HTML into the document.
        builder.InsertHtml(html);

        // Save the document to a file.
        string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // Reload the document to validate that the table and its formatting were preserved.
        Document loadedDoc = new Document(outputPath);
        Table table = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new Exception("No table was found in the loaded document.");

        // Validate the background color of the first cell (should be yellow).
        Cell firstCell = table.Rows[0].Cells[0];
        Color bgColor = firstCell.CellFormat.Shading.BackgroundPatternColor;
        if (bgColor.ToArgb() != Color.Yellow.ToArgb())
            throw new Exception("The background color of the first cell was not preserved.");

        // Validate that the left border of the first cell has a non-zero width (indicating it was imported).
        Border leftBorder = firstCell.CellFormat.Borders[BorderType.Left];
        if (leftBorder.LineWidth == 0)
            throw new Exception("The border of the first cell was not preserved.");

        // Additional validations can be added here if needed.
    }
}
