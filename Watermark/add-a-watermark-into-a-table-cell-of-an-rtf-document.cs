using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class AddWatermarkToTableCell
{
    static void Main()
    {
        // Path to the folder that contains the input RTF file.
        string dataDir = @"C:\Data\";

        // Load the existing RTF document.
        Document doc = new Document(dataDir + "input.rtf");

        // Create a DocumentBuilder to navigate and edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the first cell of the first table (row 0, column 0).
        // Adjust the indices if you need a different cell.
        builder.MoveToCell(0, 0, 0, 0);

        // Create a shape that will act as a watermark inside the cell.
        Shape watermarkShape = new Shape(doc, ShapeType.TextPlainText);
        watermarkShape.TextPath.Text = "CONFIDENTIAL";
        watermarkShape.TextPath.FontFamily = "Arial";
        // FontSize property is not available in older Aspose.Words versions; size is controlled by the shape dimensions.
        watermarkShape.Width = 300;
        watermarkShape.Height = 70;
        watermarkShape.Rotation = -40; // Diagonal appearance.
        watermarkShape.FillColor = Color.LightGray;
        watermarkShape.StrokeColor = Color.LightGray;

        // Insert the watermark shape into the current cell.
        builder.InsertNode(watermarkShape);

        // Save the modified document back to RTF format.
        doc.Save(dataDir + "output.rtf");
    }
}
