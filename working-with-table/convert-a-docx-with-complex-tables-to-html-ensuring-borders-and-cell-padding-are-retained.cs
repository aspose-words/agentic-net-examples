using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare file paths in the current working directory.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "ComplexTable.docx");
        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "ComplexTable.html");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX containing a complex table.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the outer table.
        Table outerTable = builder.StartTable();

        // Apply a thick border to the whole table.
        outerTable.SetBorder(BorderType.Left, LineStyle.Single, 2.0, System.Drawing.Color.Black, true);
        outerTable.SetBorder(BorderType.Right, LineStyle.Single, 2.0, System.Drawing.Color.Black, true);
        outerTable.SetBorder(BorderType.Top, LineStyle.Single, 2.0, System.Drawing.Color.Black, true);
        outerTable.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, System.Drawing.Color.Black, true);
        outerTable.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, System.Drawing.Color.Gray, true);
        outerTable.SetBorder(BorderType.Vertical, LineStyle.Single, 1.0, System.Drawing.Color.Gray, true);

        // Set cell padding for subsequent cells.
        builder.CellFormat.SetPaddings(10, 5, 10, 5); // left, top, right, bottom

        // First row – two cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row – first cell will contain a nested table.
        builder.InsertCell();

        // Create a nested table inside the first cell.
        Table innerTable = builder.StartTable();
        // Apply a uniform border to the nested table.
        innerTable.SetBorders(LineStyle.Single, 1.0, System.Drawing.Color.DarkBlue);
        // Adjust padding for the nested table cells.
        builder.CellFormat.SetPaddings(5, 2, 5, 2);

        builder.InsertCell();
        builder.Write("Inner 1");
        builder.InsertCell();
        builder.Write("Inner 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Inner 3");
        builder.InsertCell();
        builder.Write("Inner 4");
        builder.EndRow();

        builder.EndTable(); // End nested table.

        // Continue with the outer table's second cell.
        // The builder is currently positioned after the nested table, still inside the outer cell.
        // Move to the existing second cell of the outer row.
        builder.MoveTo(outerTable.LastRow.LastCell);
        // Write content directly into that cell (no need to insert a new cell).
        builder.Write("Regular cell");

        builder.EndRow();

        // Third row – merged cells spanning two columns.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged across columns");
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue merge.
        builder.EndRow();

        builder.EndTable(); // End outer table.

        // Save the source DOCX.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and convert it to HTML, preserving borders and padding.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export all size information (default) to keep layout identical.
            TableWidthOutputMode = HtmlElementSizeOutputMode.All
        };

        loadedDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 3. Simple validation – ensure the HTML file was created and contains border styles.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("border"))
            throw new InvalidOperationException("Converted HTML does not contain border information.");

        // Program ends without waiting for user input.
    }
}
