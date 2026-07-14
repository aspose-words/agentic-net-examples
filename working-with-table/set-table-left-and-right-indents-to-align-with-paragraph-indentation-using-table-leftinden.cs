using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph left and right indents (in points).
        builder.ParagraphFormat.LeftIndent = 30;
        builder.ParagraphFormat.RightIndent = 30;

        // Add a paragraph to demonstrate the indents.
        builder.Writeln("This paragraph has left and right indents.");

        // Start a table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Align the table's left indent with the paragraph's left indent.
        table.LeftIndent = builder.ParagraphFormat.LeftIndent;

        // Note: Aspose.Words does not provide a Table.RightIndent property,
        // so only the left indent can be aligned directly.

        // Save the document to a file.
        doc.Save("TableIndentExample.docx");
    }
}
