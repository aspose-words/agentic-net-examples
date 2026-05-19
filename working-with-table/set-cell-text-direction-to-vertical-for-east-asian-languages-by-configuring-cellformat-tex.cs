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

        // Start a table.
        Table table = builder.StartTable();

        // First cell – set text orientation to vertical for East Asian characters.
        builder.InsertCell();
        builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
        builder.Write("縦書きテキスト"); // Example Japanese vertical text.

        // Second cell – keep the default (horizontal) orientation.
        builder.InsertCell();
        builder.Write("Normal text.");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the local file system.
        const string outputPath = "CellTextDirection.docx";
        doc.Save(outputPath);
    }
}
