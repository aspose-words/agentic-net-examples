using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableToHtml
{
    public class TableHtmlConverter
    {
        /// <summary>
        /// Loads a Word document, converts the first table (including complex merged cells)
        /// to plain HTML with calculated colspan and rowspan attributes, and writes the HTML to a file.
        /// </summary>
        /// <param name="inputPath">Path to the source .doc/.docx file.</param>
        /// <param name="outputPath">Path where the resulting HTML file will be saved.</param>
        public static void ConvertTableToHtml(string inputPath, string outputPath)
        {
            // Load the document (lifecycle rule: load)
            Document doc = new Document(inputPath);

            // Get the first table in the document
            Table table = doc.FirstSection.Body.Tables[0];

            // Ensure that horizontally merged cells are represented by merge flags
            table.ConvertToHorizontallyMergedCells();

            int rowCount = table.Rows.Count;

            // Determine the maximum number of columns across all rows after conversion
            int maxColCount = 0;
            foreach (Row r in table.Rows)
                if (r.Cells.Count > maxColCount)
                    maxColCount = r.Cells.Count;

            // Track cells that are part of a previously processed merged region
            bool[,] visited = new bool[rowCount, maxColCount];

            // Build the HTML string
            StringWriter html = new StringWriter();
            html.WriteLine("<table border=\"1\" cellspacing=\"0\" cellpadding=\"5\">");

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                Row row = table.Rows[rowIndex];
                html.WriteLine("  <tr>");

                int colIndex = 0;
                while (colIndex < row.Cells.Count)
                {
                    // Skip cells that have already been covered by a previous rowspan/colspan
                    if (visited[rowIndex, colIndex])
                    {
                        colIndex++;
                        continue;
                    }

                    Cell cell = row.Cells[colIndex];
                    CellFormat fmt = cell.CellFormat;

                    // Calculate colspan
                    int colspan = 1;
                    if (fmt.HorizontalMerge == CellMerge.First)
                    {
                        int next = colIndex + 1;
                        while (next < row.Cells.Count &&
                               row.Cells[next].CellFormat.HorizontalMerge == CellMerge.Previous)
                        {
                            colspan++;
                            next++;
                        }
                    }

                    // Calculate rowspan
                    int rowspan = 1;
                    if (fmt.VerticalMerge == CellMerge.First)
                    {
                        int nextRow = rowIndex + 1;
                        while (nextRow < rowCount)
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

                    // Write the <td> element with appropriate attributes
                    html.Write("    <td");
                    if (colspan > 1)
                        html.Write($" colspan=\"{colspan}\"");
                    if (rowspan > 1)
                        html.Write($" rowspan=\"{rowspan}\"");
                    html.Write(">");

                    // Extract plain text from the cell
                    string cellText = cell.ToString(SaveFormat.Text).Trim();
                    // Encode basic HTML entities
                    cellText = System.Net.WebUtility.HtmlEncode(cellText);
                    html.Write(cellText);
                    html.WriteLine("</td>");

                    // Mark all cells covered by this merged region as visited
                    for (int r = rowIndex; r < rowIndex + rowspan; r++)
                    {
                        for (int c = colIndex; c < colIndex + colspan; c++)
                        {
                            visited[r, c] = true;
                        }
                    }

                    // Advance column index by the colspan we just processed
                    colIndex += colspan;
                }

                html.WriteLine("  </tr>");
            }

            html.WriteLine("</table>");

            // Save the generated HTML (lifecycle rule: save)
            File.WriteAllText(outputPath, html.ToString());
        }

        // Example usage
        public static void Main()
        {
            string sourceDoc = @"C:\Docs\InputWithMergedCells.docx";
            string htmlFile = @"C:\Docs\ConvertedTable.html";

            ConvertTableToHtml(sourceDoc, htmlFile);

            Console.WriteLine("HTML conversion completed.");
        }
    }
}
