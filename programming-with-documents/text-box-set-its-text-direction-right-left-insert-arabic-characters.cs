using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape with a specific size.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox = textBoxShape.TextBox;

        // Ensure the text layout flow is horizontal (default).
        textBox.LayoutFlow = LayoutFlow.Horizontal;

        // Move the builder cursor to the inside of the text box.
        builder.MoveTo(textBoxShape.LastParagraph);

        // Enable right‑to‑left formatting for the upcoming run.
        builder.Font.Bidi = true;
        // Optionally specify a right‑to‑left font and size.
        builder.Font.NameBi = "Arial";
        builder.Font.SizeBi = 24;
        builder.Font.LocaleIdBi = new CultureInfo("ar-SA").LCID;

        // Insert Arabic characters.
        builder.Write("مرحبًا بالعالم!");

        // Reset the Bidi flag if further left‑to‑right text will be added.
        builder.Font.Bidi = false;

        // Save the document to a file.
        doc.Save("TextBox_RTL_Arabic.docx");
    }
}
