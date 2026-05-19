using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Insert a table ----------
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndTable();

        // Add some spacing after the table.
        builder.Writeln();

        // ---------- Insert a text box ----------
        // Insert a text box shape of size 200x100 points.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Move the cursor inside the text box and write text.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("Text inside the box");

        // ---------- Prepare output folder ----------
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outFile = Path.Combine(outputDir, "DocumentWithTableAndTextBox.xps");

        // ---------- Save as XPS preserving layout ----------
        XpsSaveOptions saveOptions = new XpsSaveOptions
        {
            // Keep the default layout; do not apply output optimization that could alter appearance.
            OptimizeOutput = false
        };

        doc.Save(outFile, saveOptions);
    }
}
