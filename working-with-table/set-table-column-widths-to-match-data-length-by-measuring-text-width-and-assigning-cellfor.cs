using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Layout;

namespace TableColumnWidthExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a sample table with varying length text.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("ID");
            builder.InsertCell();
            builder.Write("Description");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Data rows.
            AddRow(builder, "1", "Apple", "120");
            AddRow(builder, "2", "Banana", "5");
            AddRow(builder, "3", "Cherry Pie", "12");
            AddRow(builder, "4", "Dragonfruit", "1");
            builder.EndTable();

            // Disable automatic fitting so we can set fixed column widths.
            table.AllowAutoFit = false;
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Force layout to be calculated so we can measure text widths.
            doc.UpdatePageLayout();

            // Prepare layout collector for measuring.
            LayoutCollector collector = new LayoutCollector(doc);
            LayoutEnumerator enumerator = new LayoutEnumerator(doc);

            int columnCount = table.FirstRow.Cells.Count;
            double[] maxColumnWidths = new double[columnCount];

            // Measure the width of the text in each cell and keep the maximum per column.
            foreach (Row row in table.Rows)
            {
                for (int col = 0; col < columnCount; col++)
                {
                    Cell cell = row.Cells[col];

                    // Get the first Run in the cell (there is at least one because we wrote text).
                    Run run = cell.FirstParagraph?.GetChildNodes(NodeType.Run, true)[0] as Run;
                    if (run == null)
                        continue;

                    // Get the layout entity that represents this run.
                    var entity = collector.GetEntity(run);
                    if (entity == null)
                        continue; // Safety check – prevents ArgumentNullException.

                    enumerator.Current = entity;

                    // Width is returned in points.
                    double textWidth = enumerator.Rectangle.Width;

                    // Add a small padding to avoid clipping.
                    double paddedWidth = textWidth + 5.0;
                    if (paddedWidth > maxColumnWidths[col])
                        maxColumnWidths[col] = paddedWidth;
                }
            }

            // Apply the calculated widths to each cell in the corresponding column.
            foreach (Row row in table.Rows)
            {
                for (int col = 0; col < columnCount; col++)
                {
                    Cell cell = row.Cells[col];
                    cell.CellFormat.Width = maxColumnWidths[col];
                }
            }

            // Save the document.
            string outputPath = "TableColumnWidths.docx";
            doc.Save(outputPath);
        }

        // Helper method to add a data row to the table.
        private static void AddRow(DocumentBuilder builder, string id, string description, string quantity)
        {
            builder.InsertCell();
            builder.Write(id);
            builder.InsertCell();
            builder.Write(description);
            builder.InsertCell();
            builder.Write(quantity);
            builder.EndRow();
        }
    }
}
