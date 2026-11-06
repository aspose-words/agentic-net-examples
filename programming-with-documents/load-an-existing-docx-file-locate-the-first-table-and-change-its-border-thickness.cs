using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "source.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -------------------------------------------------
        // Step 1: Create a sample DOCX file containing a table.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Save the source document.
        doc.Save(sourcePath);

        // -------------------------------------------------
        // Step 2: Load the existing DOCX file.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the first table in the document.
        Table firstTable = loadedDoc.GetChild(NodeType.Table, 0, true) as Table;
        if (firstTable != null)
        {
            // Change each border to a single line with a thickness of 2 points.
            // The last parameter (isOverrideCellBorders) is set to true to apply the border to the whole table.
            firstTable.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Black, true);
            firstTable.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Black, true);
            firstTable.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Black, true);
            firstTable.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);
        }

        // -------------------------------------------------
        // Step 3: Save the modified document.
        // -------------------------------------------------
        loadedDoc.Save(outputPath);
    }
}
