using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x2 table.
        builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Writeln("A1");
        builder.InsertCell();
        builder.Writeln("B1");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Writeln("A2");
        builder.InsertCell();
        builder.Writeln("B2");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Writeln("A3");
        builder.InsertCell();
        builder.Writeln("B3");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Merge the first two cells in the first column vertically.
        Cell firstCell = table.Rows[0].Cells[0];
        Cell secondCell = table.Rows[1].Cells[0];

        // Set vertical merge flags.
        firstCell.CellFormat.VerticalMerge = CellMerge.First;
        secondCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Update the text of the merged cell.
        firstCell.RemoveAllChildren();
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Merged A1-A2");
        para.AppendChild(run);
        firstCell.AppendChild(para);

        // Save the document.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output file was not created.");
        }
    }
}
