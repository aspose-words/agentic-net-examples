using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a caption above the table.
        // Use a SEQ field to generate automatic numbering for the label "Table".
        // The resulting caption will look like: "Table 1 Sample Table".
        builder.InsertField("SEQ Table \\* ARABIC", "1");
        builder.Write(" Sample Table");
        builder.Writeln(); // Move to the next line after the caption.

        // Build a simple 1x1 table below the caption.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = "TableCaption.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
