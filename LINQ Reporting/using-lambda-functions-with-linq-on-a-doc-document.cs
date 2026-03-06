using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // ------------------------------------------------------------
        // Example 1: Find all paragraphs that contain the word "Aspose"
        // and replace it with "Aspose.Words" using LINQ.
        // ------------------------------------------------------------
        var paragraphsWithAspose = doc.GetChildNodes(NodeType.Paragraph, true)
                                      .Cast<Paragraph>()
                                      .Where(p => p.GetText().Contains("Aspose"))
                                      .ToList();

        foreach (var paragraph in paragraphsWithAspose)
        {
            // Replace the word within the paragraph's range.
            paragraph.Range.Replace("Aspose", "Aspose.Words");
        }

        // ------------------------------------------------------------
        // Example 2: Find all bold runs and change their font color to red.
        // ------------------------------------------------------------
        var boldRuns = doc.GetChildNodes(NodeType.Run, true)
                          .Cast<Run>()
                          .Where(r => r.Font.Bold)
                          .ToList();

        foreach (var run in boldRuns)
        {
            run.Font.Color = Color.Red;
        }

        // ------------------------------------------------------------
        // Example 3: Process tables – increase numeric values in the
        // second column of each table by 10% (skip header rows).
        // ------------------------------------------------------------
        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>();

        foreach (var table in tables)
        {
            // Assume first row is a header; skip it.
            var dataRows = table.Rows.Cast<Row>().Skip(1);

            foreach (var row in dataRows)
            {
                // Get the second column (index 1).
                var cell = row.Cells[1];
                var cellText = cell.GetText().Trim();

                if (double.TryParse(cellText, out double numericValue))
                {
                    double increased = Math.Round(numericValue * 1.10, 2);

                    // Replace cell content with the new value.
                    cell.RemoveAllChildren();
                    cell.FirstParagraph.AppendChild(new Run(doc, increased.ToString()));
                }
            }
        }

        // Save the modified document using the provided Save method.
        doc.Save("Output.docx");
    }
}
