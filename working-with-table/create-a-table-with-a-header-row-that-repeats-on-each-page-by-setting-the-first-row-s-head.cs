using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableWithRepeatingHeader
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some introductory text.
            builder.Writeln("Sample document with a table that has a repeating header row.");

            // Start building the table.
            builder.StartTable();

            // ---- Header row (will be set to repeat on each page) ----
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.InsertCell();
            builder.Write("Header 3");
            builder.EndRow();

            // ---- Data rows (enough rows to span multiple pages) ----
            for (int i = 1; i <= 50; i++)
            {
                builder.InsertCell();
                builder.Write($"Row {i} Col 1");
                builder.InsertCell();
                builder.Write($"Row {i} Col 2");
                builder.InsertCell();
                builder.Write($"Row {i} Col 3");
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Retrieve the created table and set the first row's HeadingFormat to true.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Row headerRow = table.Rows[0];
            headerRow.RowFormat.HeadingFormat = true;

            // Save the document to disk.
            string outputPath = "TableWithRepeatingHeader.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("Failed to create the output document.");

            // The program ends here without waiting for user input.
        }
    }
}
