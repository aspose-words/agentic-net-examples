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

        // Start building a table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Short");
        builder.InsertCell();
        builder.Write("A much longer piece of text that should cause the column to expand.");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Another");
        builder.InsertCell();
        builder.Write("Text");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Enable AutoFit to contents so columns automatically resize to fit their content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AutoFitTable.docx");
        doc.Save(outputPath);
    }
}
