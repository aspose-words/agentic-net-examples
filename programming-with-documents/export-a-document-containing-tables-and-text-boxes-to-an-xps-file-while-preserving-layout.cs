using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string xpsPath = Path.Combine(outputDir, "DocumentWithTablesAndTextBoxes.xps");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndTable();

        // Add a paragraph break between the table and the text box.
        builder.Writeln();

        // Insert a text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 100);
        // Move the cursor inside the text box to add its content.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Font.Size = 12;
        builder.Writeln("This is a text box.");
        builder.Writeln("It contains multiple lines.");

        // Save the document to XPS format, preserving layout.
        XpsSaveOptions saveOptions = new XpsSaveOptions();
        doc.Save(xpsPath, saveOptions);

        // Optional verification that the file was created.
        if (File.Exists(xpsPath))
        {
            Console.WriteLine("XPS file saved to: " + xpsPath);
        }
    }
}
