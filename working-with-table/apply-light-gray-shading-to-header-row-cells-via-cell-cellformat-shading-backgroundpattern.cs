using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableHeaderShadingExample
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
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // ----- Data row -----
            builder.InsertCell();
            builder.Write("Data 1");
            builder.InsertCell();
            builder.Write("Data 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply light gray shading to each cell in the header row.
            foreach (Cell headerCell in table.FirstRow.Cells)
            {
                headerCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            }

            // Define output path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "HeaderShadedTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
