using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableShading
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // ----- Header row -----
            // First header cell.
            builder.InsertCell();
            builder.Write("Header 1");
            // Second header cell.
            builder.InsertCell();
            builder.Write("Header 2");
            // End the header row.
            builder.EndRow();

            // Apply light gray shading to each cell in the header row.
            foreach (Cell headerCell in table.FirstRow.Cells)
            {
                headerCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            }

            // ----- Data rows -----
            // First data row.
            builder.InsertCell();
            builder.Write("Row 1, Col 1");
            builder.InsertCell();
            builder.Write("Row 1, Col 2");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Row 2, Col 1");
            builder.InsertCell();
            builder.Write("Row 2, Col 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Define output path and ensure the directory exists.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderRowShading.docx");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
