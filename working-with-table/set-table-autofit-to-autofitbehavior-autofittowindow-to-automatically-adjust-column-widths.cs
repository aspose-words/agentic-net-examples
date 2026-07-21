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
        builder.Write("Value 1");
        builder.InsertCell();
        builder.Write("Value 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Adjust the table so its columns automatically fit the page margins.
        table.AutoFit(AutoFitBehavior.AutoFitToWindow);

        // Save the document to the current directory.
        doc.Save("TableAutoFitToWindow.docx");
    }
}
