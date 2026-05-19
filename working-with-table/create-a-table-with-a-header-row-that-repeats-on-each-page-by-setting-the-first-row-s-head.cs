using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ----- Header row (repeated on each page) -----
        builder.RowFormat.HeadingFormat = true;               // Mark this row as a heading.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.CellFormat.Width = 100;                       // Width for header cells.

        builder.InsertCell();
        builder.Write("Header Column 1");
        builder.InsertCell();
        builder.Write("Header Column 2");
        builder.EndRow();                                     // End of header row.

        // ----- Data rows (regular rows) -----
        builder.RowFormat.HeadingFormat = false;              // Subsequent rows are not headings.
        builder.CellFormat.Width = 50;                        // Width for data cells.
        builder.ParagraphFormat.ClearFormatting();

        // Add enough rows to make the table span multiple pages.
        for (int i = 0; i < 50; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Column 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Column 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithRepeatingHeader.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
