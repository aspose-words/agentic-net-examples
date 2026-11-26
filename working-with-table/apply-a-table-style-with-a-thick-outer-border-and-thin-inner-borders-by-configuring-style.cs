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
            Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

            // Create a new table style.
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

            // Configure thick outer borders.
            tableStyle.Borders[BorderType.Left].LineWidth = 2.0;
            tableStyle.Borders[BorderType.Right].LineWidth = 2.0;
            tableStyle.Borders[BorderType.Top].LineWidth = 2.0;
            tableStyle.Borders[BorderType.Bottom].LineWidth = 2.0;
            tableStyle.Borders[BorderType.Left].Color = System.Drawing.Color.Black;
            tableStyle.Borders[BorderType.Right].Color = System.Drawing.Color.Black;
            tableStyle.Borders[BorderType.Top].Color = System.Drawing.Color.Black;
            tableStyle.Borders[BorderType.Bottom].Color = System.Drawing.Color.Black;

            // Configure thin inner borders.
            tableStyle.Borders[BorderType.Horizontal].LineWidth = 0.5;
            tableStyle.Borders[BorderType.Vertical].LineWidth = 0.5;
            tableStyle.Borders[BorderType.Horizontal].Color = System.Drawing.Color.Gray;
            tableStyle.Borders[BorderType.Vertical].Color = System.Drawing.Color.Gray;

            // Apply the style to the table.
            table.Style = tableStyle;

            // Save the document.
            string outputPath = "TableStyleExample.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not created.");
        }
    }
}
