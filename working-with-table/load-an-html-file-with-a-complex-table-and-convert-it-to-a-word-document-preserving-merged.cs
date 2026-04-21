using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define file paths relative to the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string htmlPath = Path.Combine(baseDir, "sample.html");
        string outputPath = Path.Combine(baseDir, "output.docx");

        // Create a sample HTML file containing a table with both horizontal (colspan) and vertical (rowspan) merges.
        string html = @"<!DOCTYPE html>
<html>
<head><meta charset='UTF-8'></head>
<body>
<table border='1' cellspacing='0' cellpadding='5'>
  <tr>
    <td colspan='2'>Header A (colspan=2)</td>
    <td>Header B</td>
  </tr>
  <tr>
    <td rowspan='2'>Side C (rowspan=2)</td>
    <td>Data D</td>
    <td>Data E</td>
  </tr>
  <tr>
    <td>Data F</td>
    <td>Data G</td>
  </tr>
</table>
</body>
</html>";
        File.WriteAllText(htmlPath, html);

        // Load the HTML document. Aspose.Words automatically detects the format from the file extension.
        Document doc = new Document(htmlPath);

        // Ensure the document contains at least one table.
        if (doc.FirstSection?.Body?.Tables?.Count == 0)
            throw new InvalidOperationException("No table found in the loaded HTML document.");

        // Retrieve the first table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Convert cells that were merged by width (as a result of colspan/rowspan) to merge flags.
        // This allows us to inspect the merge information via CellFormat properties.
        table.ConvertToHorizontallyMergedCells();

        // ----- Validation of merged cells -----
        // First row: should have three cells with the first two horizontally merged.
        Row row0 = table.Rows[0];
        if (row0.Cells.Count != 3 ||
            row0.Cells[0].CellFormat.HorizontalMerge != CellMerge.First ||
            row0.Cells[1].CellFormat.HorizontalMerge != CellMerge.Previous ||
            row0.Cells[2].CellFormat.HorizontalMerge != CellMerge.None)
        {
            throw new InvalidOperationException("Horizontal merge in the first row is not as expected.");
        }

        // Second row: first cell should start a vertical merge.
        Row row1 = table.Rows[1];
        if (row1.Cells[0].CellFormat.VerticalMerge != CellMerge.First)
            throw new InvalidOperationException("Vertical merge start cell not detected in the second row.");

        // Third row: first cell should continue the vertical merge.
        Row row2 = table.Rows[2];
        if (row2.Cells[0].CellFormat.VerticalMerge != CellMerge.Previous)
            throw new InvalidOperationException("Vertical merge continuation cell not detected in the third row.");

        // Save the document as a Word .docx file.
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Failed to create the output Word document.", outputPath);
    }
}
