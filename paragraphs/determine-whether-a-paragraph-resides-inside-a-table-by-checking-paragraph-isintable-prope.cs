using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ParagraphInTableDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Paragraph that is NOT inside a table.
            builder.Writeln("Paragraph outside table");

            // Start a table and add a paragraph inside a cell.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Paragraph inside table");
            builder.EndTable();

            // Another paragraph outside the table.
            builder.Writeln("Another paragraph outside table");

            // Save the document (optional, just to visualize the result).
            string outputPath = "ParagraphInTable.docx";
            doc.Save(outputPath);

            // Iterate through all paragraphs in the document and report whether they are inside a table.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                // Paragraph.IsInCell is true when the paragraph is an immediate child of a Cell.
                bool isInTable = para.IsInCell;

                // Get the paragraph text without the trailing paragraph break.
                string text = para.GetText().TrimEnd('\r', '\a');

                Console.WriteLine($"\"{text}\" - Inside table: {isInTable}");
            }
        }
    }
}
