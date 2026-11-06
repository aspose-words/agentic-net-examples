using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple one‑cell table.
        builder.StartTable();
        builder.InsertCell();

        // Insert Arabic text into the cell.
        builder.Writeln("مرحبا بالعالم"); // "Hello World" in Arabic

        // Retrieve the cell that the builder is currently inside.
        Cell cell = builder.CurrentParagraph?.ParentNode as Cell;
        if (cell == null)
            throw new InvalidOperationException("Unable to locate the current cell.");

        // Set right‑to‑left (RTL) direction for all paragraphs in the cell.
        foreach (Paragraph para in cell.Paragraphs)
        {
            para.ParagraphFormat.Bidi = true; // Enables RTL layout.
        }

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        const string outputFile = "CellTextDirection.docx";
        doc.Save(outputFile);
        Console.WriteLine($"Document saved to '{outputFile}'.");
    }
}
