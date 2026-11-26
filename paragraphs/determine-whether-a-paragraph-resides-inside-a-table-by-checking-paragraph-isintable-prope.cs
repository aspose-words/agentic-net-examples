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

        // Insert a table with a single cell containing a paragraph.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Paragraph inside table.");
        builder.EndRow();
        builder.EndTable();

        // Insert a paragraph outside the table.
        builder.Writeln("Paragraph outside table.");

        // Save the document.
        doc.Save("ParagraphIsInTable.docx");

        // Iterate through all paragraphs and output whether they reside inside a table.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        int i = 0;
        foreach (Paragraph para in paragraphs)
        {
            // A paragraph is inside a table if it has a Table ancestor.
            bool isInTable = para.GetAncestor(NodeType.Table) != null;
            Console.WriteLine($"Paragraph {i}: IsInTable = {isInTable}");
            i++;
        }
    }
}
