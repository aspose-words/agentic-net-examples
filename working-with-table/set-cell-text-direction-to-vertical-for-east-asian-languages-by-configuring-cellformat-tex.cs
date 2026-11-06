using System;
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
        builder.StartTable();

        // Insert a cell and add East Asian text.
        builder.InsertCell();
        builder.Writeln("縦書きテキスト"); // Sample Japanese text.

        // Set the paragraph direction as a proxy for vertical text.
        Cell cell = builder.CurrentParagraph?.ParentNode as Cell;
        if (cell != null && cell.FirstParagraph != null)
        {
            cell.FirstParagraph.ParagraphFormat.Bidi = true;
        }

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        const string fileName = "CellTextDirection.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!System.IO.File.Exists(fileName))
            throw new Exception("The output file was not created.");
    }
}
