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

        // ---------- Insert a table into the primary header ----------
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        // Ensure there is a paragraph to host the table.
        builder.Writeln();
        // Build a 2x2 table.
        builder.StartTable();

        builder.InsertCell();
        builder.Write("Header Cell 1");
        builder.InsertCell();
        builder.Write("Header Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Header Cell 3");
        builder.InsertCell();
        builder.Write("Header Cell 4");
        builder.EndRow();

        builder.EndTable();

        // ---------- Insert a table into the primary footer ----------
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln();
        builder.StartTable();

        builder.InsertCell();
        builder.Write("Footer Cell 1");
        builder.InsertCell();
        builder.Write("Footer Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Footer Cell 3");
        builder.InsertCell();
        builder.Write("Footer Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Save the document.
        string outputPath = "HeaderFooterTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        Console.WriteLine("Document saved to: " + Path.GetFullPath(outputPath));
    }
}
