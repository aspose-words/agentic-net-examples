using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample HTML file with a complex table (colspan and rowspan).
        string htmlContent = @"<!DOCTYPE html>
<html>
<body>
<table border='1'>
<tr>
<td colspan='2'>Header</td>
</tr>
<tr>
<td rowspan='2'>RowSpan</td>
<td>Cell 1</td>
</tr>
<tr>
<td>Cell 2</td>
</tr>
</table>
</body>
</html>";

        string htmlPath = "sample.html";
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        Document doc = new Document(htmlPath);

        // Locate the first table in the document.
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
            throw new Exception("No table was found in the loaded document.");

        // Verify horizontal merge (colspan) in the first row.
        Row firstRow = table.Rows[0];
        if (firstRow.Cells.Count < 2)
            throw new Exception("Expected at least two cells in the first row.");

        Cell firstCell = firstRow.Cells[0];
        Cell secondCell = firstRow.Cells[1];
        if (firstCell.CellFormat.HorizontalMerge != CellMerge.First ||
            secondCell.CellFormat.HorizontalMerge != CellMerge.Previous)
            throw new Exception("Horizontal merge (colspan) was not preserved.");

        // Verify vertical merge (rowspan) in the first column.
        if (table.Rows.Count < 3)
            throw new Exception("Expected at least three rows in the table.");

        Row secondRow = table.Rows[1];
        Row thirdRow = table.Rows[2];
        Cell topCell = secondRow.Cells[0];
        Cell bottomCell = thirdRow.Cells[0];
        if (topCell.CellFormat.VerticalMerge != CellMerge.First ||
            bottomCell.CellFormat.VerticalMerge != CellMerge.Previous)
            throw new Exception("Vertical merge (rowspan) was not preserved.");

        // Save the document as a Word file.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The Word document was not saved correctly.");
    }
}
