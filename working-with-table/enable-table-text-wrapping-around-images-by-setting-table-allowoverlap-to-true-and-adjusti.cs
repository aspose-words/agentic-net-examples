using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A 1x1 red PNG image encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z8BQDwAF/AL+XcK6AAAAAElFTkSuQmCC";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        // Insert the image as a floating shape.
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            Shape imageShape = builder.InsertImage(imgStream);
            imageShape.WrapType = WrapType.Square;
            imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
            imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
            imageShape.Left = 0;
            imageShape.Top = 0;
            imageShape.AllowOverlap = true;
        }

        // Move the builder to a new paragraph after the image.
        builder.Writeln();

        // Start a table that will overlap the image.
        builder.StartTable();

        // Insert some cells.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Adjust layout: make the table narrower and position it over the image.
        table.PreferredWidth = PreferredWidth.FromPercent(50);
        table.LeftIndent = -20; // Move left to overlap the image.

        // Save the document.
        string outputPath = "TableWrap.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
