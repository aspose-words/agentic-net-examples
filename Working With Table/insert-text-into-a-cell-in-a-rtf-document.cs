using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class InsertTextIntoRtfCell
{
    static void Main()
    {
        // Load the existing RTF document.
        var loadOptions = new RtfLoadOptions();
        Document doc = new Document("InputDocument.rtf", loadOptions);

        // Locate the first table in the document.
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Choose the cell where the text will be inserted (e.g., first row, first column).
        Cell targetCell = table.Rows[0].Cells[0];

        // Move the DocumentBuilder cursor to the target cell.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert the desired text into the cell.
        builder.Write("Inserted text into the cell.");

        // Save the modified document back to RTF format.
        var saveOptions = new RtfSaveOptions();
        doc.Save("OutputDocument.rtf", saveOptions);
    }
}
