using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Insert a paragraph after the table.
        builder.Writeln("This is a paragraph after the table.");

        // Insert a text box.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Add text to the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Linked text box content");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Note: In newer Aspose.Words versions the Shape class provides an IsLinked property.
        // If available, you could enable linking with: textBox.IsLinked = true;
        // The current version used for compilation does not expose this property,
        // so the example creates a regular text box with content.

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify the file was created and can be reloaded.
        if (File.Exists(outputPath))
        {
            Document reloaded = new Document(outputPath);
            // Additional checks could be placed here if needed.
        }
    }
}
