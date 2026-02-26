using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class StyleSeparatorExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the first paragraph style (e.g., Heading 1) and write some text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text uses the Heading 1 style. ");

        // Insert a style separator. This creates an invisible paragraph break that
        // allows the next text to have a different paragraph style while staying
        // on the same visual line.
        builder.InsertStyleSeparator();

        // Define a custom paragraph style.
        Style customStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 10;
        customStyle.Font.Bold = false;
        customStyle.Font.Color = System.Drawing.Color.DarkGreen;

        // Apply the custom style and write additional text.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Write("This text uses a custom style.");

        // Save the document to a DOCX file.
        doc.Save("StyleSeparatorExample.docx");
    }
}
