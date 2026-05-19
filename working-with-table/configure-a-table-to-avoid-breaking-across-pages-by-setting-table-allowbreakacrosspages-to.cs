using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Populate the table with 5 rows and 3 columns.
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"Row {row + 1}, Cell {col + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Configure each row so it will not break across pages.
        foreach (Row r in table.Rows)
        {
            r.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to the local file system.
        const string outputFile = "Table_NoBreakAcrossPages.docx";
        doc.Save(outputFile);
    }
}
