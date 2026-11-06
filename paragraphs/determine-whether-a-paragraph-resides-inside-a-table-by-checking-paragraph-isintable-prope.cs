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

        // Add a paragraph that is NOT inside a table.
        builder.Writeln("This paragraph is outside any table.");

        // Start a table and add a single cell with a paragraph inside the table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("This paragraph is inside a table cell.");
        builder.EndRow();
        builder.EndTable();

        // Add another paragraph after the table.
        builder.Writeln("Another paragraph outside the table.");

        // Iterate through all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // Determine whether the paragraph resides inside a table.
            // A paragraph inside a table has a Table ancestor.
            bool isInTable = para.GetAncestor(NodeType.Table) != null;

            // Output the paragraph text and its IsInTable status.
            Console.WriteLine($"Paragraph: \"{para.GetText().TrimEnd()}\" - IsInTable: {isInTable}");
        }

        // Save the document to verify the structure.
        doc.Save("Output.docx");
    }
}
