using System;
using System.IO;
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

            // Add a paragraph outside any table.
            builder.Writeln("This paragraph is outside a table.");

            // Start a table.
            Table table = builder.StartTable();

            // First cell with a paragraph.
            builder.InsertCell();
            builder.Writeln("Paragraph inside first cell.");

            // Second cell with a paragraph.
            builder.InsertCell();
            builder.Writeln("Paragraph inside second cell.");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphInTable.docx");
            doc.Save(outputPath);

            // Iterate through all paragraphs in the document and report whether they are inside a table.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                // Paragraph.IsInCell is true when the paragraph is an immediate child of a Cell (i.e., inside a table).
                bool isInTable = para.IsInCell;
                string text = para.GetText().TrimEnd('\r', '\a').Trim();
                Console.WriteLine($"Paragraph: \"{text}\" | IsInTable: {isInTable}");
            }
        }
    }
}
