using System;
using System.IO;
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
        builder.StartTable();

        // First row.
        // Cell A (rowspan 2)
        builder.InsertCell();
        Cell cellA = (Cell)builder.CurrentParagraph.ParentNode;
        builder.Writeln("Cell A (rowspan 2)");
        cellA.CellFormat.VerticalMerge = CellMerge.First;

        // Cell B (colspan 2)
        builder.InsertCell();
        Cell cellB = (Cell)builder.CurrentParagraph.ParentNode;
        builder.Writeln("Cell B (colspan 2)");
        cellB.CellFormat.HorizontalMerge = CellMerge.First;

        // This cell will be merged horizontally with the previous one.
        builder.InsertCell();
        Cell cellC = (Cell)builder.CurrentParagraph.ParentNode;
        cellC.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.EndRow();

        // Second row.
        // Continuation of vertical merge for Cell A.
        builder.InsertCell();
        Cell cellA2 = (Cell)builder.CurrentParagraph.ParentNode;
        cellA2.CellFormat.VerticalMerge = CellMerge.Previous;

        // Cell C
        builder.InsertCell();
        Cell cellB2 = (Cell)builder.CurrentParagraph.ParentNode;
        builder.Writeln("Cell C");

        // Cell D
        builder.InsertCell();
        Cell cellC2 = (Cell)builder.CurrentParagraph.ParentNode;
        builder.Writeln("Cell D");
        builder.EndRow();

        builder.EndTable();

        // Convert the table to plain HTML with proper colspan and rowspan.
        string html = ConvertTablesToHtml(doc);

        // Save the HTML to a file.
        string outputPath = "TableOutput.html";
        File.WriteAllText(outputPath, html);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the HTML output file.");

        // Confirmation.
        Console.WriteLine($"HTML table saved to '{Path.GetFullPath(outputPath)}'.");
    }

    private static string ConvertTablesToHtml(Document doc)
    {
        var tables = doc.GetChildNodes(NodeType.Table, true);
        if (tables.Count == 0)
            return string.Empty;

        // Convert only the first table.
        Table table = (Table)tables[0];
        StringWriter writer = new StringWriter();

        writer.WriteLine("<table border=\"1\" cellspacing=\"0\" cellpadding=\"5\">");

        for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
        {
            Row row = table.Rows[rowIdx];
            writer.WriteLine("<tr>");

            for (int cellIdx = 0; cellIdx < row.Cells.Count; cellIdx++)
            {
                Cell cell = row.Cells[cellIdx];

                // Skip cells that are continuations of a merge.
                if (cell.CellFormat.HorizontalMerge == CellMerge.Previous ||
                    cell.CellFormat.VerticalMerge == CellMerge.Previous)
                    continue;

                int colspan = 1;
                int rowspan = 1;

                // Calculate colspan.
                if (cell.CellFormat.HorizontalMerge == CellMerge.First)
                {
                    for (int k = cellIdx + 1; k < row.Cells.Count; k++)
                    {
                        Cell nextCell = row.Cells[k];
                        if (nextCell.CellFormat.HorizontalMerge == CellMerge.Previous)
                            colspan++;
                        else
                            break;
                    }
                }

                // Calculate rowspan.
                if (cell.CellFormat.VerticalMerge == CellMerge.First)
                {
                    for (int r = rowIdx + 1; r < table.Rows.Count; r++)
                    {
                        Row nextRow = table.Rows[r];
                        if (cellIdx < nextRow.Cells.Count)
                        {
                            Cell belowCell = nextRow.Cells[cellIdx];
                            if (belowCell.CellFormat.VerticalMerge == CellMerge.Previous)
                                rowspan++;
                            else
                                break;
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                // Get the plain text of the cell.
                string cellText = cell.GetText().TrimEnd('\a', '\r', '\n');

                // Build the <td> element.
                writer.Write("<td");
                if (colspan > 1)
                    writer.Write($" colspan=\"{colspan}\"");
                if (rowspan > 1)
                    writer.Write($" rowspan=\"{rowspan}\"");
                writer.Write($">{WebUtility.HtmlEncode(cellText)}</td>");
            }

            writer.WriteLine("</tr>");
        }

        writer.WriteLine("</table>");
        return writer.ToString();
    }
}
