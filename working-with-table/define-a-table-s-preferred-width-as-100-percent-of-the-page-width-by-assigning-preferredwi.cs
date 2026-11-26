using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TablePreferredWidthExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple one‑cell table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Sample cell");
            builder.EndRow();
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
                throw new InvalidOperationException("Table was not created.");

            // Set the table's preferred width to 100 % of the page width.
            // PreferredWidth does not have a two‑argument constructor in the current API.
            // Use the static factory method FromPercent to create the required value.
            table.PreferredWidth = PreferredWidth.FromPercent(100);

            // Save the document.
            string outputPath = "TablePreferredWidth.docx";
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
