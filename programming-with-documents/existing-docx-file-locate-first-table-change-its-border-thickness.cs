using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file if it exists; otherwise create a new document with a simple table.
        Document doc;
        const string inputPath = "Input.docx";
        const string outputPath = "Output.docx";

        if (File.Exists(inputPath))
        {
            doc = new Document(inputPath);
        }
        else
        {
            doc = new Document();
            Section section = (Section)doc.AppendChild(new Section(doc));
            Body body = new Body(doc);
            section.AppendChild(body);

            // Create a simple 1x1 table.
            Table table = new Table(doc);
            Row row = new Row(doc);
            Cell cell = new Cell(doc);
            cell.AppendChild(new Paragraph(doc));
            row.AppendChild(cell);
            table.AppendChild(row);
            body.AppendChild(table);
        }

        // Ensure there is at least one table in the document.
        Table targetTable = null;
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            targetTable = doc.FirstSection.Body.Tables[0];
        }

        if (targetTable != null)
        {
            // Change the thickness of all four outline borders.
            // The last parameter (true) removes any explicit cell borders.
            targetTable.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Black, true);
            targetTable.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Black, true);
            targetTable.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Black, true);
            targetTable.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);
        }
        else
        {
            Console.WriteLine("No table found in the document.");
        }

        // Save the modified document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
