using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAllowBreakAcrossPagesExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3x3 table.
            builder.StartTable();

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Writeln($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Retrieve the created table.
            Table table = doc.GetChild(NodeType.Table, 0, true) as Table;
            if (table == null)
                throw new InvalidOperationException("Table was not created.");

            // Configure each row to avoid breaking across pages.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.AllowBreakAcrossPages = false;
            }

            // Save the document.
            string outputPath = "TableAllowBreakAcrossPages.docx";
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }
    }
}
