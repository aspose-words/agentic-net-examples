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

        // Insert a paragraph outside any table.
        builder.Writeln("Paragraph outside a table.");

        // Start a table.
        Table table = builder.StartTable();

        // Insert first cell with a paragraph.
        builder.InsertCell();
        builder.Writeln("Paragraph inside first cell.");

        // Insert second cell with a paragraph.
        builder.InsertCell();
        builder.Writeln("Paragraph inside second cell.");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("Output.docx");

        // Iterate through all paragraphs and report whether they are inside a table.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // A paragraph is inside a table if it has a Table ancestor.
            bool isInTable = para.GetAncestor(NodeType.Table) != null;
            Console.WriteLine($"Paragraph text: \"{para.GetText().TrimEnd()}\" - IsInTable: {isInTable}");
        }
    }
}
