using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 3x2 table.
            builder.StartTable();

            // Row 1
            builder.InsertCell();
            builder.Write("Row 1, Col 1");
            builder.InsertCell();
            builder.Write("Row 1, Col 2");
            builder.EndRow();

            // Row 2
            builder.InsertCell();
            builder.Write("Row 2, Col 1");
            builder.InsertCell();
            builder.Write("Row 2, Col 2");
            builder.EndRow();

            // Row 3
            builder.InsertCell();
            builder.Write("Row 3, Col 1");
            builder.InsertCell();
            builder.Write("Row 3, Col 2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Merge cells vertically in the first column (index 0) across the three rows.
            Cell firstCell = table.Rows[0].Cells[0];
            Cell secondCell = table.Rows[1].Cells[0];
            Cell thirdCell = table.Rows[2].Cells[0];

            // Set vertical merge flags.
            firstCell.CellFormat.VerticalMerge = CellMerge.First;      // Start of merged region
            secondCell.CellFormat.VerticalMerge = CellMerge.Previous; // Continuation
            thirdCell.CellFormat.VerticalMerge = CellMerge.Previous;  // Continuation

            // Save the document.
            string outputPath = "MergedTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output file was not created.");
        }
    }
}
