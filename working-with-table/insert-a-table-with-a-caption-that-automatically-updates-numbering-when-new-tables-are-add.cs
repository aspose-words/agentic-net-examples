using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("First table, cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("First table, cell 2");
        builder.EndRow();
        builder.EndTable();

        // Caption for the first table with automatic numbering.
        InsertTableCaption(builder, "First table description");

        // Second table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Second table, cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Second table, cell 2");
        builder.EndRow();
        builder.EndTable();

        // Caption for the second table.
        InsertTableCaption(builder, "Second table description");

        // Save the document.
        string outputPath = "TableCaption.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }

    private static void InsertTableCaption(DocumentBuilder builder, string description)
    {
        // Move to a new paragraph after the table.
        builder.Writeln();

        // Write the static part of the caption.
        builder.Write("Table ");

        // Insert a SEQ field that automatically numbers tables.
        // The field will be updated when the document is opened in Word.
        builder.InsertField("SEQ Table \\* ARABIC");

        // Write the rest of the caption.
        builder.Write($": {description}");
        builder.Writeln();
    }
}
