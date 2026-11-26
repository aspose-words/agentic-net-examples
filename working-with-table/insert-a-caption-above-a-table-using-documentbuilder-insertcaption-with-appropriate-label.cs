using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a caption above the table with label "Table" and automatic numbering.
        // Since InsertCaption is not available, we build the caption manually.
        builder.Write("Table ");
        // Insert a SEQ field that will generate the table number.
        builder.InsertField("SEQ Table \\* ARABIC", "1");
        builder.Write(". Sample data");
        builder.Writeln(); // Move to a new paragraph after the caption.

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Save the document.
        string outputPath = "TableWithCaption.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to {Path.GetFullPath(outputPath)}");
    }
}
