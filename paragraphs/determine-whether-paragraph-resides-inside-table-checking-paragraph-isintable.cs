using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with one cell containing a paragraph.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Paragraph inside a table cell");
        builder.EndTable();

        // Add a paragraph that is not inside a table.
        builder.Writeln("Paragraph outside any table");

        // Iterate over all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // The IsInCell property is true if the paragraph is an immediate child of a Cell.
            bool isInTable = para.IsInCell;

            // Output the paragraph text and its IsInCell status.
            Console.WriteLine($"Text: \"{para.GetText().Trim()}\"  IsInCell: {isInTable}");
        }

        // Save the document to verify the content.
        doc.Save("IsInCellDemo.docx");
    }
}
