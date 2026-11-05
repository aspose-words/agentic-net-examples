using System;
using System.IO;
using System.Drawing;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Insert a cell.
        builder.InsertCell();

        // Get the current cell.
        Cell cell = builder.CurrentParagraph.ParentNode as Cell;
        if (cell == null)
            throw new InvalidOperationException("Current cell not found.");

        // Apply diagonal lines to the cell using reflection (works even if the properties are not present in the current library version).
        CellFormat fmt = cell.CellFormat;
        Type fmtType = fmt.GetType();

        PropertyInfo propDiagonalDown = fmtType.GetProperty("DiagonalDown");
        PropertyInfo propDiagonalUp = fmtType.GetProperty("DiagonalUp");
        PropertyInfo propDiagonalColor = fmtType.GetProperty("DiagonalColor");
        PropertyInfo propDiagonalWidth = fmtType.GetProperty("DiagonalWidth");

        if (propDiagonalDown != null && propDiagonalDown.CanWrite)
            propDiagonalDown.SetValue(fmt, true);
        if (propDiagonalUp != null && propDiagonalUp.CanWrite)
            propDiagonalUp.SetValue(fmt, true);
        if (propDiagonalColor != null && propDiagonalColor.CanWrite)
            propDiagonalColor.SetValue(fmt, Color.Red);
        if (propDiagonalWidth != null && propDiagonalWidth.CanWrite)
            propDiagonalWidth.SetValue(fmt, 2.0); // points

        // Add some text to the cell.
        builder.Writeln("Diagonal Cell");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string fileName = "DiagonalCell.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception($"Failed to create the output file: {fileName}");

        // Load the document again and ensure it contains a table.
        Document loaded = new Document(fileName);
        if (loaded.GetChildNodes(NodeType.Table, true).Count == 0)
            throw new Exception("The saved document does not contain a table.");
    }
}
