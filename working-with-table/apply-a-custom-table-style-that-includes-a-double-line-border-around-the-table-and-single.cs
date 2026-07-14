using System;
using System.IO;
using System.Drawing;               // Added for Color
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 3x3 table with sample content.
            Table table = builder.StartTable();

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Apply double line borders to the outer edges of the table.
            // The last parameter 'true' removes any existing explicit cell borders for those sides.
            table.SetBorder(BorderType.Left,   LineStyle.Double, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Right,  LineStyle.Double, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Top,    LineStyle.Double, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Double, 2.0, Color.Black, true);

            // Apply single line borders to all inner cell edges.
            foreach (Row r in table.Rows)
            {
                foreach (Cell c in r.Cells)
                {
                    c.CellFormat.Borders.LineStyle = LineStyle.Single;
                    c.CellFormat.Borders.LineWidth = 1.0;
                    c.CellFormat.Borders.Color = Color.Black;
                }
            }

            // Save the document to the local file system.
            string outputPath = "CustomTableStyle.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
