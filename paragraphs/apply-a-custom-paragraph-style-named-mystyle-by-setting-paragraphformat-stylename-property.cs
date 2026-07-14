using System;
using System.Drawing;
using System.IO;
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
        builder.ParagraphFormat.StyleName = "MyStyle";
        builder.Writeln("This paragraph is formatted with the custom style MyStyle.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MyStyleParagraph.docx");
        doc.Save(outputPath);
    }
}
