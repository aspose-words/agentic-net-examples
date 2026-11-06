using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom paragraph style named "MyStyle".
        Style myStyle = doc.Styles.Add(StyleType.Paragraph, "MyStyle");
        myStyle.Font.Name = "Arial";
        myStyle.Font.Size = 14;
        myStyle.Font.Color = Color.Blue;

        // Use DocumentBuilder to insert a paragraph and apply the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Apply the custom style by setting the StyleName property.
        builder.ParagraphFormat.StyleName = "MyStyle";
        builder.Writeln("This paragraph uses the custom style \"MyStyle\".");

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
