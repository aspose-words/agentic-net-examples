using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // Define a custom character style named "MyCustomStyle".
        // ------------------------------------------------------------
        Style customStyle = doc.Styles.Add(StyleType.Character, "MyCustomStyle");
        customStyle.Font.Name = "Arial";          // Font name
        customStyle.Font.Size = 14;               // Font size (points)
        customStyle.Font.Color = Color.Blue;     // Font color

        // Write a paragraph using the default document style.
        builder.Writeln("This paragraph uses the default formatting.");

        // Apply the custom style to the next text run.
        builder.Font.StyleName = "MyCustomStyle";
        builder.Writeln("This paragraph is formatted with MyCustomStyle.");

        // Optional: clear the style to return to default formatting.
        builder.Font.ClearFormatting();

        // Save the document to disk.
        doc.Save("StyledText.docx");
    }
}
