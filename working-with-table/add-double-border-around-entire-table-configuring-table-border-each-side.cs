using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Remove any existing borders to start with a clean slate.
        table.ClearBorders();

        // Apply a double line border to each side of the table.
        // Parameters: BorderType, LineStyle, line width (points), color, override cell borders.
        table.SetBorder(BorderType.Left,   LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Right,  LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Top,    LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 1.5, Color.Black, true);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
