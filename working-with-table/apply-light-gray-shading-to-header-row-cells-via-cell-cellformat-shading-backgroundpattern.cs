using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin building a table.
        builder.StartTable();

        // Create a header row with three cells.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Header {i + 1}");

            // Apply light gray shading to the current header cell.
            Cell headerCell = (Cell)builder.CurrentParagraph.ParentNode;
            headerCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        }
        // End the header row.
        builder.EndRow();

        // Add a simple data row for illustration.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Data {i + 1}");
        }
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "TableWithHeaderShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
