using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;   // Needed for VerticalAlignment enum

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1x1 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Set the table's vertical alignment on the page to the middle.
        // Use RelativeVerticalAlignment for floating tables.
        table.RelativeVerticalAlignment = VerticalAlignment.Center;

        // Save the document.
        const string outputFile = "TableVerticalAlignment.docx";
        doc.Save(outputFile);
    }
}
