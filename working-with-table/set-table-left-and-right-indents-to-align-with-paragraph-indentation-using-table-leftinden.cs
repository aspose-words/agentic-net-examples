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

        // Set paragraph indentation (left and right) to 20 points.
        builder.ParagraphFormat.LeftIndent = 20;
        builder.ParagraphFormat.RightIndent = 20;
        builder.Writeln("This paragraph has left and right indents of 20 points.");

        // Start a new table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Align the table's left indent with the paragraph's left indent.
        table.LeftIndent = builder.ParagraphFormat.LeftIndent;

        // Note: The Aspose.Words API does not provide a Table.RightIndent property,
        // so only the left indent can be set directly.

        // Save the document to disk.
        string outputPath = "TableIndent.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output file was not created.");
        }

        // Program completed successfully.
        Console.WriteLine("Document saved to " + Path.GetFullPath(outputPath));
    }
}
