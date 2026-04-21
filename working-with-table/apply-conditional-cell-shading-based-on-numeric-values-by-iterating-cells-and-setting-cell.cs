using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableShading
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with numeric values.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Data rows.
            AddDataRow(builder, "Apples", "30");
            AddDataRow(builder, "Bananas", "75");
            AddDataRow(builder, "Cherries", "120");
            builder.EndTable();

            // Iterate through all cells and apply shading based on the numeric value.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tables)
            {
                foreach (Row row in tbl.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        // Extract the cell text.
                        string cellText = cell.ToString(SaveFormat.Text).Trim();

                        // Try to parse a numeric value.
                        if (double.TryParse(cellText, out double number))
                        {
                            // Apply shading: values > 50 get light green, otherwise light gray.
                            if (number > 50)
                                cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
                            else
                                cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                        }
                    }
                }
            }

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConditionalShadingTable.docx");
            doc.Save(outputPath);

            // Validate that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }

        // Helper method to add a data row to the table.
        private static void AddDataRow(DocumentBuilder builder, string item, string quantity)
        {
            builder.InsertCell();
            builder.Write(item);
            builder.InsertCell();
            builder.Write(quantity);
            builder.EndRow();
        }
    }
}
