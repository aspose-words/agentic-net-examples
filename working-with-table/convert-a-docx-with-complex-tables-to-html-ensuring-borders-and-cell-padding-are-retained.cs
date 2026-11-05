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
        // Paths for the generated files.
        string docPath = "ComplexTable.docx";
        string htmlPath = "ComplexTable.html";

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        Cell cell = (Cell)builder.CurrentParagraph.ParentNode;
        SetCellPadding(cell, 5);
        builder.Writeln("Header 1");

        // First row, second cell.
        builder.InsertCell();
        cell = (Cell)builder.CurrentParagraph.ParentNode;
        SetCellPadding(cell, 5);
        builder.Writeln("Header 2");

        // End first row.
        builder.EndRow();

        // Second row, first cell (merged across two columns).
        builder.InsertCell();
        cell = (Cell)builder.CurrentParagraph.ParentNode;
        SetCellPadding(cell, 5);
        cell.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Writeln("Merged cell spanning two columns");

        // Second row, second cell (merged continuation).
        builder.InsertCell();
        cell = (Cell)builder.CurrentParagraph.ParentNode;
        cell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End second row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table to apply borders.
        Table table = doc.FirstSection.Body.Tables[0];
        table.ClearBorders();
        table.SetBorder(BorderType.Left, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Vertical, LineStyle.Single, 1.0, Color.Black, true);

        // Save the DOCX file.
        doc.Save(docPath);

        // Load the document and convert it to HTML.
        Document htmlDoc = new Document(docPath);
        htmlDoc.Save(htmlPath, SaveFormat.Html);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
        {
            throw new Exception("HTML conversion failed: output file not found.");
        }

        Console.WriteLine("Conversion completed. HTML saved at: " + Path.GetFullPath(htmlPath));
    }

    private static void SetCellPadding(Cell cell, double padding)
    {
        cell.CellFormat.LeftPadding = padding;
        cell.CellFormat.RightPadding = padding;
        cell.CellFormat.TopPadding = padding;
        cell.CellFormat.BottomPadding = padding;
    }
}
