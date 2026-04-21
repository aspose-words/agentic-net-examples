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
        // Create a sample document with a complex merged‑cell table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // First row – a cell merged horizontally (2 columns) and vertically (2 rows).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("A");

        builder.InsertCell(); // horizontally merged part
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write(string.Empty);

        builder.InsertCell(); // normal cell
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("B");
        builder.EndRow();

        // Second row – continuation of the vertical merge.
        builder.InsertCell(); // vertical merge continuation
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        builder.Write(string.Empty);

        builder.InsertCell(); // normal cell
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("C");

        builder.InsertCell(); // normal cell
        builder.Write("D");
        builder.EndRow();

        // Third row – regular cells.
        builder.InsertCell();
        builder.Write("E");
        builder.InsertCell();
        builder.Write("F");
        builder.InsertCell();
        builder.Write("G");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Ensure the table uses merge flags rather than width‑based merging.
        table.ConvertToHorizontallyMergedCells();

        // Convert the table to plain HTML with proper colspan/rowspan.
        string html = ConvertTableToHtml(table);

        // Save the HTML to a file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedTable.html");
        File.WriteAllText(outputPath, html, Encoding.UTF8);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the HTML output file.");

        // Optionally, write a confirmation to the console (no user interaction required).
        Console.WriteLine($"HTML file saved to: {outputPath}");
    }

    private static string ConvertTableToHtml(Table table)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("<table border=\"1\" cellspacing=\"0\" cellpadding=\"5\">");

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            Row row = table.Rows[rowIndex];
            sb.AppendLine("  <tr>");

            for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
            {
                Cell cell = row.Cells[colIndex];

                // Skip cells that are merged from the left or above.
                if (cell.CellFormat.HorizontalMerge == CellMerge.Previous ||
                    cell.CellFormat.VerticalMerge == CellMerge.Previous)
                    continue;

                // Determine colspan.
                int colspan = 1;
                if (cell.CellFormat.HorizontalMerge == CellMerge.First)
                {
                    int next = colIndex + 1;
                    while (next < row.Cells.Count &&
                           row.Cells[next].CellFormat.HorizontalMerge == CellMerge.Previous)
                    {
                        colspan++;
                        next++;
                    }
                }

                // Determine rowspan.
                int rowspan = 1;
                if (cell.CellFormat.VerticalMerge == CellMerge.First)
                {
                    int nextRow = rowIndex + 1;
                    while (nextRow < table.Rows.Count)
                    {
                        Cell belowCell = table.Rows[nextRow].Cells[colIndex];
                        if (belowCell.CellFormat.VerticalMerge == CellMerge.Previous)
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

                // Get the cell text and HTML‑encode it.
                string cellText = WebUtility.HtmlEncode(cell.GetText().Trim());

                // Build the <td> element with appropriate attributes.
                sb.Append("    <td");
                if (colspan > 1)
                    sb.Append($" colspan=\"{colspan}\"");
                if (rowspan > 1)
                    sb.Append($" rowspan=\"{rowspan}\"");
                sb.Append($">{cellText}</td>");
                sb.AppendLine();
            }

            sb.AppendLine("  </tr>");
        }

        sb.AppendLine("</table>");
        return sb.ToString();
    }
}
