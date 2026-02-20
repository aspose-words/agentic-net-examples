using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace InsertTableIntoDocm
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new Word document (DOCM format)
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table
            Table table = builder.StartTable();

            // Insert first row with two cells (header row)
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Insert second row
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Insert third row
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // Finish the table
            builder.EndTable();

            // Optionally apply auto‑fit to contents
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document as a macro‑enabled file (DOCM)
            doc.Save("OutputDocument.docm", SaveFormat.Docm);
        }
    }
}
