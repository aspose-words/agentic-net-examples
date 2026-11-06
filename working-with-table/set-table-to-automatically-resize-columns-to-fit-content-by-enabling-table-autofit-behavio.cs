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

        // Build a simple table with some sample content.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Header 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("Header 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Short");

        // Second row, second cell with longer text.
        builder.InsertCell();
        builder.Writeln("This is a much longer piece of text that should cause the column to expand.");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Enable AutoFit to contents so columns resize automatically.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        string outputPath = "TableAutoFit.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
