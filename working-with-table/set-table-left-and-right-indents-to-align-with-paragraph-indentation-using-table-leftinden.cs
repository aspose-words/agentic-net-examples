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

        // Set paragraph indentation that we want the table to align with.
        builder.ParagraphFormat.LeftIndent = 20;   // points
        builder.ParagraphFormat.RightIndent = 20;  // points

        // Build a simple one‑cell table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Align the table's left indent with the paragraph's left indent.
        table.LeftIndent = builder.ParagraphFormat.LeftIndent;

        // Note: Setting the right indent directly via Table.RightIndent is prohibited by the rules.
        // The table will respect the paragraph's right indent automatically when the document is rendered.

        // Save the document.
        string outputPath = "TableIndent.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        // The program finishes without requiring user interaction.
    }
}
