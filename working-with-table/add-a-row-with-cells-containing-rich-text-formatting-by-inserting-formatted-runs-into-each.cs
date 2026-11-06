using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build an initial table with a single header row.
        builder.StartTable();

        builder.InsertCell();
        builder.Write("Header 1");

        builder.InsertCell();
        builder.Write("Header 2");

        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Create a new row to add to the table.
        Row newRow = new Row(doc);
        table.Rows.Add(newRow);

        // Define the number of cells (columns) in the new row.
        int columnCount = 2;

        for (int col = 0; col < columnCount; col++)
        {
            // Create a new cell and add it to the row.
            Cell cell = new Cell(doc);
            newRow.Cells.Add(cell);

            // Create a paragraph inside the cell.
            Paragraph para = new Paragraph(doc);
            cell.AppendChild(para);

            // Add a bold run.
            Run boldRun = new Run(doc, "Bold");
            boldRun.Font.Bold = true;
            para.AppendChild(boldRun);

            // Add an italic run (preceded by a space for separation).
            Run italicRun = new Run(doc, " Italic");
            italicRun.Font.Italic = true;
            para.AppendChild(italicRun);

            // Add a colored run (preceded by a space for separation).
            Run colorRun = new Run(doc, " Red");
            colorRun.Font.Color = Color.Red;
            para.AppendChild(colorRun);
        }

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
