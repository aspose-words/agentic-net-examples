using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableAlignment
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table and keep a reference to it.
            Table table = builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Write("Cell 2");

            // End the first row.
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Write("Cell 3");

            // Second row, second cell.
            builder.InsertCell();
            builder.Write("Cell 4");

            // End the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Align the table to the center of the page.
            table.Alignment = TableAlignment.Center;

            // Prepare output folder.
            string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Save the document.
            string outputPath = Path.Combine(artifactsDir, "CenteredTable.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
