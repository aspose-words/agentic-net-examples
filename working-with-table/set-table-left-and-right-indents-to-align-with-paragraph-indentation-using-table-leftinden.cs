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

        // Set paragraph indentation that we want the table to match.
        builder.ParagraphFormat.LeftIndent = 30;   // points
        builder.ParagraphFormat.RightIndent = 30; // points

        // Add a paragraph with some text.
        builder.Writeln("This paragraph has left and right indents of 30 points.");

        // Start a table.
        builder.StartTable();

        // Insert a single cell with some text.
        builder.InsertCell();
        builder.Writeln("Table cell content.");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Align the table's left indent with the paragraph's left indent.
        table.LeftIndent = builder.ParagraphFormat.LeftIndent;

        // Note: Table.RightIndent is not used due to API restrictions.

        // Save the document.
        string outputPath = "TableIndent.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output file was not created.");
        }
    }
}
