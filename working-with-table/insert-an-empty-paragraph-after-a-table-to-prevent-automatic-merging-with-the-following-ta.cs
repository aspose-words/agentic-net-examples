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

        // ---------- First table ----------
        Table table1 = builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Table1 R1C1");
        builder.InsertCell();
        builder.Write("Table1 R1C2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Table1 R2C1");
        builder.InsertCell();
        builder.Write("Table1 R2C2");
        builder.EndRow();

        builder.EndTable();

        // Insert an empty paragraph to prevent the next table from merging with the previous one.
        builder.InsertParagraph(); // This creates a blank paragraph.

        // ---------- Second table ----------
        Table table2 = builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Table2 R1C1");
        builder.InsertCell();
        builder.Write("Table2 R1C2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Table2 R2C1");
        builder.InsertCell();
        builder.Write("Table2 R2C2");
        builder.EndRow();

        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablesWithSeparator.docx");
        doc.Save(outputPath);
    }
}
