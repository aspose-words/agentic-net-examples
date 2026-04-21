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

        // Start a table.
        Table table = builder.StartTable();

        // First row.
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

        // Finish the table.
        builder.EndTable();

        // Adjust the table to fit the window (page margins).
        table.AutoFit(AutoFitBehavior.AutoFitToWindow);

        // Save the document.
        const string outputPath = "AutoFitToWindow.docx";
        doc.Save(outputPath);
    }
}
