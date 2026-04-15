using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class ExportToXps
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Insert a table ----------
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table Cell 1");
        builder.InsertCell();
        builder.Write("Table Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Add some spacing after the table.
        builder.Writeln();
        builder.Writeln();

        // ---------- Insert a text box ----------
        // Create a text box shape of size 200x100 points.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Move the cursor inside the text box and add text.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("This is a text box.");

        // ---------- Save the document as XPS ----------
        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithTableAndTextBox.xps");

        // Create XpsSaveOptions (default options preserve layout).
        XpsSaveOptions saveOptions = new XpsSaveOptions();

        // Save the document to XPS format.
        doc.Save(outputPath, saveOptions);
    }
}
