using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoPdfCell
{
    static void Main()
    {
        // Path where the PDF will be saved.
        string outputPath = @"C:\Temp\CellInPdf.pdf";

        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a simple 1x1 table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();
        builder.InsertCell();                     // First (and only) cell.
        builder.Write("Initial placeholder text"); // Optional placeholder.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the cell we just created.
        Cell cell = table.Rows[0].Cells[0];

        // Remove any existing content from the cell.
        cell.RemoveAllChildren();

        // Create a new paragraph that will hold the desired text.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Hello, this text is inside a PDF table cell!");
        paragraph.AppendChild(run);

        // Add the paragraph to the cell.
        cell.AppendChild(paragraph);

        // Save the document as a PDF file.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
