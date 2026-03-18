using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text.
        builder.Writeln("Sample text with red fill and 30% transparency.");

        // Apply a solid red fill to the font.
        builder.Font.Fill.Solid(Color.Red);

        // Set the fill transparency to 30% (0.3).
        builder.Font.Fill.Transparency = 0.3;

        // Save the document.
        doc.Save("TextFillRedTransparent.docx");
    }
}
