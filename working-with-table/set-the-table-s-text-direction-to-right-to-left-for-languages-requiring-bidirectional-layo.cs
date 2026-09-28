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

        // Start a table.
        builder.StartTable();

        // First row, first cell with Arabic text.
        builder.InsertCell();
        builder.Writeln("مرحبا بالعالم");
        // Set paragraph direction to right‑to‑left.
        builder.CurrentParagraph.ParagraphFormat.Bidi = true;

        // First row, second cell with Hebrew text.
        builder.InsertCell();
        builder.Writeln("שלום עולם");
        builder.CurrentParagraph.ParagraphFormat.Bidi = true;

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("نص إضافي");
        builder.CurrentParagraph.ParagraphFormat.Bidi = true;

        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("טקסט נוסף");
        builder.CurrentParagraph.ParagraphFormat.Bidi = true;

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Ensure every paragraph inside the table is set to right‑to‑left.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    foreach (Paragraph para in cell.Paragraphs)
                    {
                        para.ParagraphFormat.Bidi = true;
                    }
                }
            }
        }

        // Save the document.
        string outputPath = "TableRightToLeft.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
