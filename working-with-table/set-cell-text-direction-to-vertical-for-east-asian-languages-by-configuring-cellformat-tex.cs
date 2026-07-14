using System;
using System.IO;
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

        // First cell – set text direction to vertical for East Asian languages.
        builder.InsertCell();
        builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
        builder.Write("縦書きテキスト"); // Sample East Asian vertical text.

        // Second cell – normal horizontal text.
        builder.InsertCell();
        builder.CellFormat.Orientation = TextOrientation.Horizontal;
        builder.Write("Horizontal text");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "CellVerticalText.docx");
        doc.Save(outputPath);
    }
}
