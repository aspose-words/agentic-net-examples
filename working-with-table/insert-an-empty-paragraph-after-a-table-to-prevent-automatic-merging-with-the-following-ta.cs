using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- First table ----------
            builder.StartTable();
            builder.InsertCell();
            builder.Write("First table, cell 1");
            builder.InsertCell();
            builder.Write("First table, cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("First table, cell 3");
            builder.InsertCell();
            builder.Write("First table, cell 4");
            builder.EndTable();

            // Insert an empty paragraph to separate the tables.
            // Writeln() without arguments creates a paragraph break with no text.
            builder.Writeln();

            // ---------- Second table ----------
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Second table, cell 1");
            builder.InsertCell();
            builder.Write("Second table, cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Second table, cell 3");
            builder.InsertCell();
            builder.Write("Second table, cell 4");
            builder.EndTable();

            // Save the document.
            const string outputPath = "Result.docx";
            doc.Save(outputPath);

            // Simple validation: ensure there are exactly two tables and an empty paragraph between them.
            Table[] tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToArray();

            if (tables.Length != 2)
                throw new InvalidOperationException("The document does not contain the expected number of tables.");

            // The node immediately after the first table should be a Paragraph with no text.
            Node nodeAfterFirstTable = tables[0].NextSibling;
            if (nodeAfterFirstTable == null ||
                nodeAfterFirstTable.NodeType != NodeType.Paragraph ||
                !string.IsNullOrWhiteSpace(nodeAfterFirstTable.GetText()))
            {
                throw new InvalidOperationException("The empty paragraph separating the tables was not inserted correctly.");
            }

            // Execution finished successfully.
        }
    }
}
