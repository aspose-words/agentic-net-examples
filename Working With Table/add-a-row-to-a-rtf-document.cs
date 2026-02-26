using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToRtf
{
    static void Main()
    {
        // Paths to the source RTF document and the result document.
        string inputPath = "input.rtf";
        string outputPath = "output.rtf";

        // Load the existing RTF document.
        Document doc = new Document(inputPath);

        // Assume the document contains at least one table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row belonging to the same document.
        Row newRow = new Row(doc);

        // Append the new row to the end of the table.
        table.AppendChild(newRow);

        // Ensure the row has at least one cell.
        newRow.EnsureMinimum();

        // Add some text to the first cell of the new row.
        newRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "New row content"));

        // Save the modified document as RTF.
        doc.Save(outputPath, SaveFormat.Rtf);
    }
}
