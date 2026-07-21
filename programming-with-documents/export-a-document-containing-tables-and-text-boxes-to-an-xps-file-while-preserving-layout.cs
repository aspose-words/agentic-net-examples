using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Insert a table ----------
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

        // ---------- Insert a text box ----------
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;
        textBox.Height = 100;
        textBox.WrapType = WrapType.None;
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        textBox.Left = 100;
        textBox.Top = 300;

        // Add text inside the text box.
        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "Text inside the text box");
        tbParagraph.AppendChild(tbRun);
        textBox.AppendChild(tbParagraph);

        // Place the text box into the document.
        builder.InsertNode(textBox);

        // ---------- Save as XPS ----------
        string xpsPath = Path.Combine(outputDir, "DocumentWithTableAndTextBox.xps");
        XpsSaveOptions saveOptions = new XpsSaveOptions
        {
            // Preserve the original layout (default behavior). Explicitly disable output optimization.
            OptimizeOutput = false
        };
        doc.Save(xpsPath, saveOptions);
    }
}
