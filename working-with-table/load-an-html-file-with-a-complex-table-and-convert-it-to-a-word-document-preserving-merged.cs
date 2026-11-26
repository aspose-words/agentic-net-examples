using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a sample HTML file with a complex table that contains merged cells.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head><meta charset='UTF-8'></head>
<body>
<table border='1' style='border-collapse:collapse;'>
    <tr>
        <td colspan='2' style='background:#D3D3D3; text-align:center;'><b>Header Merged Across Two Columns</b></td>
        <td>Header 3</td>
    </tr>
    <tr>
        <td rowspan='2' style='background:#ADD8E6; text-align:center;'><b>Row Span 2</b></td>
        <td>Cell 2,2</td>
        <td>Cell 2,3</td>
    </tr>
    <tr>
        <td colspan='2' style='background:#90EE90; text-align:center;'><b>Col Span 2</b></td>
    </tr>
    <tr>
        <td>Cell 4,1</td>
        <td>Cell 4,2</td>
        <td>Cell 4,3</td>
    </tr>
</table>
</body>
</html>";

        // Write the HTML content to a temporary file.
        string tempHtmlPath = Path.Combine(Path.GetTempPath(), "sample_table.html");
        File.WriteAllText(tempHtmlPath, htmlContent);

        // Load the HTML file into an Aspose.Words Document.
        LoadOptions loadOptions = new HtmlLoadOptions();
        Document doc = new Document(tempHtmlPath, loadOptions);

        // Save the document as a Word file.
        string outputDocPath = Path.Combine(Directory.GetCurrentDirectory(), "ConvertedTable.docx");
        doc.Save(outputDocPath);

        // Validate that the output file was created.
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The output Word document was not created.");

        // Validate that merged cells were preserved.
        // Look for at least one cell with a horizontal or vertical merge setting.
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
            throw new InvalidOperationException("No table was found in the loaded document.");

        bool hasMergedCell = false;
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                if (cell.CellFormat.HorizontalMerge != CellMerge.None ||
                    cell.CellFormat.VerticalMerge != CellMerge.None)
                {
                    hasMergedCell = true;
                    break;
                }
            }
            if (hasMergedCell) break;
        }

        if (!hasMergedCell)
            throw new InvalidOperationException("Merged cells were not preserved during conversion.");

        // Clean up temporary HTML file.
        File.Delete(tempHtmlPath);
    }
}
