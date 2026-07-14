using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class ExportToXps
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

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
        builder.InsertCell();
        builder.Write("Table Cell 3");
        builder.InsertCell();
        builder.Write("Table Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Add some spacing between elements.
        builder.Writeln();
        builder.Writeln();

        // ---------- Insert a text box ----------
        // Create a text box shape of size 200x100 points.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Move the cursor inside the text box and add text.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("This is a text box.");

        // ---------- Save as XPS ----------
        string xpsPath = Path.Combine(outputDir, "DocumentWithTableAndTextBox.xps");
        XpsSaveOptions saveOptions = new XpsSaveOptions();
        // Preserve layout (default behavior). Optionally, enable high‑quality rendering.
        saveOptions.UseHighQualityRendering = true;

        doc.Save(xpsPath, saveOptions);

        // Simple verification that the file was created.
        if (File.Exists(xpsPath))
        {
            Console.WriteLine("XPS file created successfully at: " + xpsPath);
        }
        else
        {
            Console.WriteLine("Failed to create XPS file.");
        }
    }
}
