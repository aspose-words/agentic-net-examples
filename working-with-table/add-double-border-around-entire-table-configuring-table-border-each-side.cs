using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        const string inputPath = "Input.docx";
        const string outputPath = "Output.docx";

        Document doc;
        if (File.Exists(inputPath))
        {
            doc = new Document(inputPath);
        }
        else
        {
            // Create a simple document with one table if the input file does not exist.
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Sample table created because Input.docx was not found.");
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable();
        }

        // Ensure the document has at least one table.
        Table tableToFormat = null;
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            tableToFormat = doc.FirstSection.Body.Tables[0];
        }
        else
        {
            // If no tables exist, create one.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Added table for border demonstration:");
            tableToFormat = builder.StartTable();
            builder.InsertCell();
            builder.Write("A1");
            builder.InsertCell();
            builder.Write("B1");
            builder.EndRow();
            builder.EndTable();
        }

        // Remove any existing borders.
        tableToFormat.ClearBorders();

        // Apply a double line border to each side of the table.
        tableToFormat.SetBorder(BorderType.Left,   LineStyle.Double, 1.5, Color.Black, true);
        tableToFormat.SetBorder(BorderType.Right,  LineStyle.Double, 1.5, Color.Black, true);
        tableToFormat.SetBorder(BorderType.Top,    LineStyle.Double, 1.5, Color.Black, true);
        tableToFormat.SetBorder(BorderType.Bottom, LineStyle.Double, 1.5, Color.Black, true);

        // Save the modified document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
