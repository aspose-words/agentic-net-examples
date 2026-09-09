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

        // Build a simple 2x1 table with two paragraphs inside its cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Paragraph inside first cell.");
        builder.InsertCell();
        builder.Writeln("Paragraph inside second cell.");
        builder.EndRow();
        builder.EndTable();

        // Add a paragraph that is not inside a table.
        builder.Writeln("Paragraph outside any table.");

        // Iterate through all paragraphs in the document and report whether they are inside a table cell.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // The IsInCell property is true when the paragraph is an immediate child of a Cell.
            bool isInTable = para.IsInCell;
            string text = para.GetText().TrimEnd('\r', '\a'); // Remove paragraph break characters.
            Console.WriteLine($"Text: \"{text}\" | IsInCell: {isInTable}");
        }

        // Save the document to verify the structure.
        doc.Save("ParagraphIsInCell.docx");
    }
}
