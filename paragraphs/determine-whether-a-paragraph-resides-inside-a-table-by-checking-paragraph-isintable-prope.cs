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
        builder.Writeln("Paragraph outside any table.");

        // Build a simple 1‑row, 2‑cell table.
        builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Writeln("Paragraph inside first cell.");

        // Second cell.
        builder.InsertCell();
        builder.Writeln("Paragraph inside second cell.");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Iterate through all paragraphs in the document and report whether each is inside a table cell.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // Trim the paragraph text to remove the trailing paragraph break.
            string text = para.GetText().TrimEnd('\r', '\a');
            Console.WriteLine($"\"{text}\" - IsInCell: {para.IsInCell}");
        }

        // Save the document to the local file system.
        doc.Save("ParagraphInTable.docx");
    }
}
