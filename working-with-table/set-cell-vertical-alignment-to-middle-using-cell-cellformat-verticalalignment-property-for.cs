using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // First cell – set vertical alignment to middle.
        builder.InsertCell();
        builder.Writeln("Middle aligned");
        Cell firstCell = (Cell)builder.CurrentParagraph.ParentNode;
        firstCell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

        // Second cell – default alignment.
        builder.InsertCell();
        builder.Writeln("Normal");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string filePath = "VerticalAlignmentTable.docx";
        doc.Save(filePath);

        // Verify that the file was created.
        if (!File.Exists(filePath))
            throw new Exception($"Failed to create the output file: {filePath}");

        // Reload the document and verify the vertical alignment.
        Document loadedDoc = new Document(filePath);
        Table table = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
        Cell loadedFirstCell = table.Rows[0].Cells[0];
        if (loadedFirstCell.CellFormat.VerticalAlignment != CellVerticalAlignment.Center)
            throw new Exception("The cell vertical alignment was not set to middle as expected.");

        // Indicate successful completion (optional).
        Console.WriteLine("Document created and vertical alignment verified successfully.");
    }
}
