using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom paragraph style named "MyStyle".
        Style myStyle = doc.Styles.Add(StyleType.Paragraph, "MyStyle");
        // Example formatting for the custom style.
        myStyle.Font.Name = "Arial";
        myStyle.Font.Size = 14;
        myStyle.Font.Bold = true;

        // Insert a paragraph using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will use the custom style.");

        // Retrieve the paragraph that was just added.
        Paragraph paragraph = doc.LastSection.Body.Paragraphs[doc.LastSection.Body.Paragraphs.Count - 1];

        // Apply the custom style by setting the StyleName.
        paragraph.ParagraphFormat.StyleName = "MyStyle";

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
