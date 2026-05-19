using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample DOCX and the resulting HTML.
        string docxPath = Path.Combine(outputDir, "ComplexTable.docx");
        string htmlPath = Path.Combine(outputDir, "ComplexTable.html");

        // ---------- Create a sample document with a complex table ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the outer table.
        Table table = builder.StartTable();

        // Apply a uniform border to the whole outer table.
        table.SetBorder(BorderType.Left, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Vertical, LineStyle.Single, 1.0, Color.Black, true);

        // Set cell padding for all subsequent cells.
        builder.CellFormat.SetPaddings(5, 5, 5, 5);

        // First row – three cells, merge the last two horizontally.
        builder.InsertCell();
        builder.Write("Header 1");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First; // Start merge.
        builder.Write("Merged Header");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue merge.
        // No text needed for merged part.

        builder.EndRow();

        // Reset merge for next rows.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // Second row – normal cells.
        builder.InsertCell();
        builder.Write("Row 1, Col 1");

        builder.InsertCell();
        builder.Write("Row 1, Col 2");

        builder.InsertCell();
        builder.Write("Row 1, Col 3");

        builder.EndRow();

        // Third row – a nested table inside the first cell.
        builder.InsertCell();

        // Start nested table.
        Table nested = builder.StartTable();

        // Apply borders to the nested table (using DotDash which is supported).
        nested.SetBorder(BorderType.Left, LineStyle.DotDash, 0.5, Color.Gray, true);
        nested.SetBorder(BorderType.Right, LineStyle.DotDash, 0.5, Color.Gray, true);
        nested.SetBorder(BorderType.Top, LineStyle.DotDash, 0.5, Color.Gray, true);
        nested.SetBorder(BorderType.Bottom, LineStyle.DotDash, 0.5, Color.Gray, true);
        nested.SetBorder(BorderType.Horizontal, LineStyle.DotDash, 0.5, Color.Gray, true);
        nested.SetBorder(BorderType.Vertical, LineStyle.DotDash, 0.5, Color.Gray, true);

        // Set padding for cells of the nested table.
        builder.CellFormat.SetPaddings(2, 2, 2, 2);

        // First row of nested table.
        builder.InsertCell();
        builder.Write("Nested 1");
        builder.InsertCell();
        builder.Write("Nested 2");
        builder.EndRow();

        // End nested table.
        builder.EndTable();

        // Continue the outer table's third row.
        builder.InsertCell();
        builder.Write("Row 2, Col 2");

        builder.InsertCell();
        builder.Write("Row 2, Col 3");

        builder.EndRow();

        // Finish the outer table.
        builder.EndTable();

        // Save the sample DOCX.
        doc.Save(docxPath);

        // ---------- Load the sample document and convert it to HTML ----------
        Document loadDoc = new Document(docxPath);

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Preserve all size information.
            TableWidthOutputMode = HtmlElementSizeOutputMode.All
        };

        loadDoc.Save(htmlPath, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new Exception("HTML conversion failed: output file not found.");

        // Indicate success.
        Console.WriteLine("Conversion completed successfully.");
    }
}
