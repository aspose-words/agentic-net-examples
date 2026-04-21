using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build first sample table.
            Table table1 = builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Build second sample table.
            Table table2 = builder.StartTable();
            builder.InsertCell();
            builder.Write("A");
            builder.InsertCell();
            builder.Write("B");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("C");
            builder.InsertCell();
            builder.Write("D");
            builder.EndRow();
            builder.EndTable();

            // Iterate over all tables in the document and set their style to "Table Grid".
            foreach (Table tbl in doc.GetChildNodes(NodeType.Table, true))
            {
                tbl.StyleIdentifier = StyleIdentifier.TableGrid;
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");
            doc.Save(outputPath);
        }
    }
}
