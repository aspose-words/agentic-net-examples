using System.Drawing;
using Aspose.Words;

class FontConsistencyExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to set up a common font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Set the desired font size (in points) and color.
        builder.Font.Size = 14;               // 14‑point font
        builder.Font.Color = Color.Blue;      // Blue color

        // The builder's current font settings will be applied to all text it inserts.
        // Insert the first run of text.
        builder.Writeln("This is the first run with consistent formatting.");

        // Insert the second run of text; it inherits the same font settings.
        builder.Writeln("This is the second run sharing the same font.");

        // Insert a third run to demonstrate further consistency.
        builder.Writeln("Third run also uses the same font size and color.");

        // Save the document to a file.
        doc.Save("FontConsistency.docx");
    }
}
