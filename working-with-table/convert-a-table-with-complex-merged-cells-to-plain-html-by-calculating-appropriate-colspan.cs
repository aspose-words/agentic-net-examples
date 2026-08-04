using System;
using System.IO;
using System.Text;
using System.Net;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with complex merged cells.
        // Row 1: first cell spans two columns (colspan), third cell normal.
        // Row 2: first cell normal, second cell starts a vertical merge (rowspan), third cell normal.
        // Row 3: first cell normal, second cell continues the vertical merge, third cell normal.
        Table table = builder.StartTable();

        // ----- Row 1 -----
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header (colspan 2)");

        builder.InsertCell(); // continuation of horizontal merge
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text for merged part.

        builder.InsertCell(); // normal cell
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Header 3");
        builder.EndRow();

        // ----- Row 2 -----
        builder.InsertCell(); // normal cell
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row2 Col1");

        builder.InsertCell(); // start of vertical merge
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged (rowspan 2)");

        builder.InsertCell(); // normal cell
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row2 Col3");
        builder.EndRow();

        // ----- Row 3 -----
        builder.InsertCell(); // normal cell
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row3 Col1");

        builder.InsertCell(); // continuation of vertical merge
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text for merged part.

        builder.InsertCell(); // normal cell
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row3 Col3");
        builder.EndRow();

        builder.EndTable();

        // Ensure horizontal merges are represented by merge flags.
        table.ConvertToHorizontallyMergedCells();

        // Convert the table to plain HTML with proper colspan and rowspan.
        string html = ConvertTableToHtml(table);

        // Save the HTML to a file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableExport.html");
        File.WriteAllText(outputPath, html, Encoding.UTF8);
    }

    private static string ConvertTableToHtml(Table table)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("<table border=\"1\" cellspacing=\"0\" cellpadding=\"5\">");

        int rowCount = table.Rows.Count;

        for (int r = 0; r < rowCount; r++)
        {
            sb.AppendLine("  <tr>");
            Row row = table.Rows[r];
            for (int i = 0; i < row.Cells.Count; i++)
            {
                Cell cell = row.Cells[i];

                // Skip cells that are continuations of a merged region.
                if (cell.CellFormat.HorizontalMerge == CellMerge.Previous ||
                    cell.CellFormat.VerticalMerge == CellMerge.Previous)
                    continue;

                // Determine colspan.
                int colspan = 1;
                if (cell.CellFormat.HorizontalMerge == CellMerge.First)
                {
                    int j = i + 1;
                    while (j < row.Cells.Count && row.Cells[j].CellFormat.HorizontalMerge == CellMerge.Previous)
                    {
                        colspan++;
                        j++;
                    }
                    // Advance the loop index to the last merged cell.
                    i = j - 1;
                }

                // Determine rowspan.
                int rowspan = 1;
                if (cell.CellFormat.VerticalMerge == CellMerge.First)
                {
                    int nextRow = r + 1;
                    while (nextRow < rowCount)
                    {
                        // The column index in the next row that aligns with the current cell.
                        // For this simple example we assume column positions are consistent across rows.
                        Cell nextCell = table.Rows[nextRow].Cells[i];
                        if (nextCell.CellFormat.VerticalMerge == CellMerge.Previous)
                        {
                            rowspan++;
                            nextRow++;
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                // Get cell text and HTML‑encode it.
                string cellText = WebUtility.HtmlEncode(cell.GetText().Trim());

                // Build the <td> element.
                sb.Append("    <td");
                if (colspan > 1) sb.Append($" colspan=\"{colspan}\"");
                if (rowspan > 1) sb.Append($" rowspan=\"{rowspan}\"");
                sb.Append($">{cellText}</td>");
                sb.AppendLine();
            }
            sb.AppendLine("  </tr>");
        }

        sb.AppendLine("</table>");
        return sb.ToString();
    }
}
