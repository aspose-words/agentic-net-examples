using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertDrawingCanvas
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle that will act as a drawing canvas.
        // The shape is positioned relative to the page (left:100pt, top:100pt) 
        // and sized to 300pt width x 200pt height. No text wrapping.
        Shape canvas = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            300, 200,
            WrapType.None);

        // Optional: give the canvas a name and visual styling.
        canvas.Name = "MyDrawingCanvas";
        canvas.Stroke.Color = Color.Black;   // border color
        canvas.Fill.Color = Color.White;     // background color

        // Save the document using DML compliance so the shape is stored correctly.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("DrawingCanvas.docx", saveOptions);
    }
}
