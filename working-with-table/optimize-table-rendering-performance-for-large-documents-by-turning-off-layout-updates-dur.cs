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

        // Build a large table (500 rows × 10 columns) while layout updates are not required.
        // The layout will be calculated only once after the table is fully constructed.
        Table table = builder.StartTable();

        int rowCount = 500;
        int columnCount = 10;

        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < columnCount; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Perform a single layout pass now that all modifications are complete.
        doc.UpdatePageLayout();

        // Save the resulting document.
        doc.Save("LargeTable.docx");
    }
}
