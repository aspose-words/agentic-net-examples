using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Mark the first row as a heading row that repeats on each page.
        builder.RowFormat.HeadingFormat = true;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.CellFormat.Width = 100;

        // Header cells.
        builder.InsertCell();
        builder.Write("ID");
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Reset formatting for the data rows.
        builder.RowFormat.HeadingFormat = false;
        builder.ParagraphFormat.ClearFormatting();
        builder.CellFormat.ClearFormatting();

        // Add sample data rows.
        for (int i = 1; i <= 30; i++)
        {
            builder.InsertCell();
            builder.Write(i.ToString());

            builder.InsertCell();
            builder.Write($"Item {i}");

            builder.InsertCell();
            builder.Write((i * 10).ToString());

            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableWithHeaders.docx";
        doc.Save(outputPath);
    }
}
