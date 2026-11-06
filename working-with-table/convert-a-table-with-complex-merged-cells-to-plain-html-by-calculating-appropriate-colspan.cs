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

        // Build a table with horizontal and vertical merges.
        builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header 1-2");
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.Write(""); // merged cell placeholder
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Header 3");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertical 1-2");
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("R2C2");
        builder.InsertCell();
        builder.Write("R2C3");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        builder.Write(""); // merged cell placeholder
        builder.InsertCell();
        builder.Write("R3C2");
        builder.InsertCell();
        builder.Write("R3C3");
        builder.EndRow();

        builder.EndTable();

        // Save the sample document (optional, for verification).
        string docPath = "sample.docx";
        doc.Save(docPath);

        // Convert the first table to plain HTML with proper colspan and rowspan.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        StringBuilder html = new StringBuilder();
        html.AppendLine("<table border=\"1\" cellspacing=\"0\" cellpadding=\"5\">");

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            Row row = table.Rows[rowIndex];
            html.AppendLine("<tr>");

            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                Cell cell = row.Cells[cellIndex];

                // Skip cells that are part of a previous horizontal or vertical merge.
                if (cell.CellFormat.HorizontalMerge == CellMerge.Previous ||
                    cell.CellFormat.VerticalMerge == CellMerge.Previous)
                {
                    continue;
                }

                // Determine colspan.
                int colspan = 1;
                if (cell.CellFormat.HorizontalMerge == CellMerge.First)
                {
                    int next = cellIndex + 1;
                    while (next < row.Cells.Count && row.Cells[next].CellFormat.HorizontalMerge == CellMerge.Previous)
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
                        Row followingRow = table.Rows[nextRow];
                        if (cellIndex < followingRow.Cells.Count &&
                            followingRow.Cells[cellIndex].CellFormat.VerticalMerge == CellMerge.Previous)
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

                // Build the <td> element.
                html.Append("<td");
                if (colspan > 1) html.Append($" colspan=\"{colspan}\"");
                if (rowspan > 1) html.Append($" rowspan=\"{rowspan}\"");
                html.Append(">");

                // Extract and HTML‑encode the cell text.
                string cellText = cell.GetText().TrimEnd('\a');
                html.Append(WebUtility.HtmlEncode(cellText));

                html.AppendLine("</td>");
            }

            html.AppendLine("</tr>");
        }

        html.AppendLine("</table>");

        // Save the generated HTML.
        string htmlPath = "table.html";
        File.WriteAllText(htmlPath, html.ToString(), Encoding.UTF8);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
        {
            throw new InvalidOperationException("HTML output file was not created.");
        }
    }
}
