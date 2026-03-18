using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a single cell so the table is valid.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample text");
        builder.EndRow();
        builder.EndTable();

        // 1 centimeter = 28.35 points (72 points per inch, 2.54 cm per inch).
        double oneCentimeterInPoints = 28.35;

        // Set left indent of the table to one centimeter.
        table.LeftIndent = oneCentimeterInPoints;

        // NOTE: Aspose.Words for .NET versions prior to 23.5 do NOT expose a RightIndent property.
        // If you need a right margin, either:
        //   1) Upgrade to a version that provides Table.RightIndent and then add:
        //        table.RightIndent = oneCentimeterInPoints;
        //   2) Or simulate a right margin by reducing the table's preferred width:
        //        table.PreferredWidth = PreferredWidth.FromPercent(100 - (oneCentimeterInPoints * 2 / doc.PageSetup.PageWidth * 100));
        // The example below uses the first option (commented) for clarity.
        // table.RightIndent = oneCentimeterInPoints; // Uncomment after upgrading.

        // Save the document.
        doc.Save("TableWithMargins.docx");
    }
}
