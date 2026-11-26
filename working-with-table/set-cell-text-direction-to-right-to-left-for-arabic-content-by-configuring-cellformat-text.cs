using System;
using System.IO;
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

        // Insert a cell and write Arabic text.
        builder.InsertCell();
        builder.Writeln("مرحبا بالعالم"); // Arabic for "Hello World"

        // Retrieve the cell that was just created.
        Cell cell = (Cell)builder.CurrentParagraph.ParentNode;

        // Set the paragraph direction to right‑to‑left.
        Paragraph para = cell.FirstParagraph;
        para.ParagraphFormat.Bidi = true;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string fileName = "CellTextDirection.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception("The output file was not created.");

        // Reload the document and confirm the text direction.
        Document loadedDoc = new Document(fileName);
        Table table = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
        Cell loadedCell = table.Rows[0].Cells[0];
        bool isBidi = loadedCell.FirstParagraph.ParagraphFormat.Bidi;
        if (!isBidi)
            throw new Exception("The cell text direction was not set to right‑to‑left.");
    }
}
