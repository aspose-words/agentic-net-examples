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

        // Use DocumentBuilder to construct a simple 2‑column table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("First column");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Second column");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply AutoFit to make the table width adjust to the page margins.
        table.AutoFit(AutoFitBehavior.AutoFitToWindow);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AutoFitTable.docx");
        doc.Save(outputPath);
    }
}
