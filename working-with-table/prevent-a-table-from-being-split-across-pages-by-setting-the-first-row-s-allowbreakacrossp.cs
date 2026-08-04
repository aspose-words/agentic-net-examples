using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3‑row, 2‑column table.
            Table table = builder.StartTable();

            // First row (header).
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Third row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Prevent the first row from being split across pages.
            Row firstRow = table.FirstRow;
            firstRow.RowFormat.AllowBreakAcrossPages = false;

            // Ensure the output directory exists.
            string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Save the document.
            string outputPath = Path.Combine(artifactsDir, "Table_NoBreakAcrossPages.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not created.");

            // Reload the document and confirm the setting persisted.
            Document loadedDoc = new Document(outputPath);
            Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
            bool allowBreak = loadedTable.FirstRow.RowFormat.AllowBreakAcrossPages;
            if (allowBreak)
                throw new InvalidOperationException("AllowBreakAcrossPages was not set to false.");

            // Example completed.
        }
    }
}
