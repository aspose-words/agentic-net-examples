using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Apply a text watermark to the document using the Watermark class.
        doc.Watermark.SetText("Sample Watermark");

        // Build a table with a first row that will display the watermark behind its cells.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();

        // First row – three cells with header text.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Header {i + 1}");
        }
        builder.EndRow();

        // Second row – sample data.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Data {i + 1}");
        }
        builder.EndRow();

        builder.EndTable();

        // Save the resulting document.
        doc.Save("WatermarkedTable.docx");
    }
}
