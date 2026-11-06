using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Apply borders: top and bottom solid, left and right none.
        table.SetBorder(BorderType.Top,    LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Left,   LineStyle.None,   0.0, Color.Black, false);
        table.SetBorder(BorderType.Right,  LineStyle.None,   0.0, Color.Black, false);

        // Save the document.
        string outputPath = "TableBorders.docx";
        doc.Save(outputPath);

        // Verify the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully: {Path.GetFullPath(outputPath)}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
