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

        // First row with two cells.
        builder.InsertCell();
        builder.Write("Short");
        builder.InsertCell();
        builder.Write("A much longer piece of text that should cause the column to expand.");
        builder.EndRow();

        // Second row with two cells.
        builder.InsertCell();
        builder.Write("Another short");
        builder.InsertCell();
        builder.Write("More long content to test auto fit behavior.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Enable automatic column resizing to fit the cell contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the local file system.
        const string outputPath = "AutoFitTable.docx";
        doc.Save(outputPath);
    }
}
