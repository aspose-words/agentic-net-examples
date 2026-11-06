using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Move to a new paragraph after the table.
        builder.Writeln();

        // Add a caption paragraph styled as "Caption" and bookmark it.
        builder.StartBookmark("TableCaption");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
        builder.Write("Table ");
        builder.InsertField("SEQ Table \\* ARABIC");
        builder.Write(": Sample Table");
        builder.EndBookmark("TableCaption");
        builder.Writeln();

        // Insert a reference to the caption.
        builder.Writeln("Reference to the table:");
        builder.InsertField("REF TableCaption \\h");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        else
        {
            Console.WriteLine("Failed to create the document.");
        }
    }
}
