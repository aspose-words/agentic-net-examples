using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Apply a solid shading texture and set the background color to light gray.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Insert a cell and add some text.
        builder.InsertCell();
        builder.Writeln("Cell with light gray shading");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "CellShading.docx");
        doc.Save(outputPath);
    }
}
