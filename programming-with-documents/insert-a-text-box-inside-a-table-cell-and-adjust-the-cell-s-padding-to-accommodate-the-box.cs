using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a single cell.
        builder.StartTable();
        builder.InsertCell();
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table and its first cell.
        Table table = doc.FirstSection.Body.Tables[0];
        Cell cell = table.FirstRow.FirstCell;

        // Adjust cell padding so the text box fits comfortably.
        cell.CellFormat.TopPadding = 10;
        cell.CellFormat.BottomPadding = 10;
        cell.CellFormat.LeftPadding = 10;
        cell.CellFormat.RightPadding = 10;

        // Create a text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 150;
        textBox.Height = 80;
        textBox.WrapType = WrapType.Inline; // Ensure it stays inside the cell.

        // Add some text to the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Text inside the box");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Insert the text box into the cell.
        cell.FirstParagraph.AppendChild(textBox);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
        }
        else
        {
            Console.WriteLine("Failed to create the document.");
        }
    }
}
