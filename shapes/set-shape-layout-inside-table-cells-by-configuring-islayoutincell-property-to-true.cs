using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert a single cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Move the cursor to the first paragraph of the first cell.
        builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph);

        // Insert a floating rectangle shape inside the cell.
        // The shape is positioned relative to the left and top margins of the page.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 50,
            RelativeVerticalPosition.TopMargin, 100,
            100, 100,
            WrapType.None);

        // Configure the shape to be displayed inside the table cell.
        shape.IsLayoutInCell = true;
        // The IsLayoutInCell property works only for floating shapes.
        shape.WrapType = WrapType.None;

        // End the table.
        builder.EndTable();

        // Prepare the output directory and file path.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ShapeLayoutInCell.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
