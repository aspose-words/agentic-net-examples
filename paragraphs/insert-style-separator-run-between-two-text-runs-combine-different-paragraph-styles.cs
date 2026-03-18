using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the first part of the line using a built‑in paragraph style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text is in a Heading 1 style. ");

        // Insert a style separator so the next text can have a different paragraph style
        // while staying on the same visual line.
        builder.InsertStyleSeparator();

        // Create a custom paragraph style.
        Style customStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 10;
        customStyle.Font.Bold = false;

        // Apply the custom style and write the second part of the line.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Write("This text is in a custom style.");

        // Save the resulting document.
        doc.Save("StyleSeparatorExample.docx");
    }
}
