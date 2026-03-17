using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle with a single snipped corner.
        // Width = 200 points, Height = 100 points.
        Shape shape = builder.InsertShape(ShapeType.SingleCornerSnipped, 200, 100);

        // Apply a light gray fill to the shape.
        shape.FillColor = Color.LightGray;

        // Set the snip amount (corner radius) via the shape's adjustment value.
        // The Adjustments collection is read‑only, but each Adjustment object is mutable.
        // For a SingleCornerSnipped shape there is exactly one adjustment.
        shape.Adjustments[0].Value = 10; // 10 points radius

        // Save the document using a compliance level that preserves DML shapes.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("SnipCornerRectangle.docx", saveOptions);
    }
}
