using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a caption above the table: "Table 1 Above".
        // Use a field to generate the table number automatically.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
        builder.InsertField("SEQ Table \\* ARABIC", "1");
        builder.Write(" Above");
        builder.Writeln(); // Move to the next paragraph.

        // Build a simple 2x2 table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string fileName = "TableWithCaption.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception("The output document was not created.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(fileName)}");
    }
}
