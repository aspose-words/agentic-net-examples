using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableBordersExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a single row with two cells.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell with thick top border");
            builder.InsertCell();
            builder.Write("Cell with thin bottom border");
            builder.EndRow();

            // Apply a thick top border (4 points) and a thin bottom border (1 point).
            // The 'true' flag overrides any existing cell borders.
            table.SetBorder(BorderType.Top, LineStyle.Single, 4.0, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);

            // Finish the table.
            builder.EndTable();

            // Save the document to a file.
            string outputPath = "TableBorders.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
