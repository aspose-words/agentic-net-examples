using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace DeleteTableByKeyword
{
    public class Program
    {
        public static void Main()
        {
            // Output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeletedTable.docx");

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First table – does NOT contain the keyword.
            Table table1 = builder.StartTable();
            builder.InsertCell();
            builder.Write("First table, cell 1.");
            builder.InsertCell();
            builder.Write("First table, cell 2.");
            builder.EndRow();
            builder.EndTable();

            // Second table – contains the keyword "DeleteMe".
            Table table2 = builder.StartTable();
            builder.InsertCell();
            builder.Write("This table will be deleted. Keyword: DeleteMe");
            builder.InsertCell();
            builder.Write("Another cell.");
            builder.EndRow();
            builder.EndTable();

            // Third table – does NOT contain the keyword.
            Table table3 = builder.StartTable();
            builder.InsertCell();
            builder.Write("Third table, cell 1.");
            builder.InsertCell();
            builder.Write("Third table, cell 2.");
            builder.EndRow();
            builder.EndTable();

            // Keyword to search for.
            const string keyword = "DeleteMe";

            // Get all tables in the document. Use LINQ to create a snapshot array
            // because we will modify the document while iterating.
            Table[] allTables = doc.GetChildNodes(NodeType.Table, true)
                                   .OfType<Table>()
                                   .ToArray();

            foreach (Table tbl in allTables)
            {
                // If the table's text contains the keyword, remove the table.
                if (tbl.Range.Text.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    tbl.Remove();
                }
            }

            // Save the resulting document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully: {outputPath}");
            }
            else
            {
                throw new InvalidOperationException("Failed to save the document.");
            }
        }
    }
}
