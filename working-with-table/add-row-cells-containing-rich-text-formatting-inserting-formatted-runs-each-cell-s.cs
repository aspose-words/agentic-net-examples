using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class AddFormattedRowExample
{
    static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table (if the document already has a table, this part can be omitted).
        Table table = builder.StartTable();

        // -----------------------------------------------------------------
        // Existing rows (optional) – here we add a simple header row.
        // -----------------------------------------------------------------
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // -----------------------------------------------------------------
        // Add a new row where each cell contains rich‑text formatting.
        // -----------------------------------------------------------------
        // Begin the new row.
        builder.InsertCell(); // First cell of the new row.

        // Get the cell that was just created.
        Cell firstCell = (Cell)builder.CurrentParagraph.ParentNode;

        // Clear any existing content (the cell already contains an empty paragraph).
        firstCell.FirstParagraph.Runs.Clear();

        // Create a bold run.
        Run boldRun = new Run(doc, "Bold Text");
        boldRun.Font.Bold = true;
        firstCell.FirstParagraph.AppendChild(boldRun);

        // Create a red, italic run.
        Run redItalicRun = new Run(doc, " Red Italic");
        redItalicRun.Font.Italic = true;
        redItalicRun.Font.Color = Color.Red; // Fixed property name
        firstCell.FirstParagraph.AppendChild(redItalicRun);

        // -----------------------------------------------------------------
        // Second cell with different formatting.
        // -----------------------------------------------------------------
        builder.InsertCell(); // Second cell of the same row.
        Cell secondCell = (Cell)builder.CurrentParagraph.ParentNode;
        secondCell.FirstParagraph.Runs.Clear();

        // Underlined, blue run.
        Run underlineRun = new Run(doc, "Underlined");
        underlineRun.Font.Underline = Underline.Single;
        underlineRun.Font.Color = Color.Blue; // Fixed property name
        secondCell.FirstParagraph.AppendChild(underlineRun);

        // Append a normal run after the formatted one.
        Run normalRun = new Run(doc, " Normal");
        secondCell.FirstParagraph.AppendChild(normalRun);

        // Finish the row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        doc.Save("FormattedRow.docx");
    }
}
