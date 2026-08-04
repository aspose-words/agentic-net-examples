using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the styled template document.
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        // Define a custom paragraph style named "MyStyle" with specific formatting.
        Style templateStyle = templateDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
        templateStyle.Font.Name = "Arial";
        templateStyle.Font.Size = 16;
        templateStyle.Font.Color = Color.Blue;

        // Apply the custom style to a heading in the template.
        templateBuilder.ParagraphFormat.StyleName = "MyStyle";
        templateBuilder.Writeln("Template Heading");

        // Add some normal text after the heading.
        templateBuilder.ParagraphFormat.StyleName = "Normal";
        templateBuilder.Writeln("This is the template content before insertion.");

        // Create the source document to be inserted.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // Define a style with the same name but different formatting to demonstrate style clash handling.
        Style sourceStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
        sourceStyle.Font.Name = "Times New Roman";
        sourceStyle.Font.Size = 14;
        sourceStyle.Font.Color = Color.Red;

        // Apply the source style to some text.
        sourceBuilder.ParagraphFormat.StyleName = "MyStyle";
        sourceBuilder.Writeln("Inserted Heading from Source Document");

        // Add additional content in the source document.
        sourceBuilder.ParagraphFormat.StyleName = "Normal";
        sourceBuilder.Writeln("This is the content from the source document.");

        // Insert the source document into the template using UseDestinationStyles mode.
        templateBuilder.MoveToDocumentEnd();
        templateBuilder.InsertDocument(sourceDoc, ImportFormatMode.UseDestinationStyles);

        // Save the merged document as HTML.
        string outputPath = "MergedDocument.html";
        templateDoc.Save(outputPath, SaveFormat.Html);

        // Optional: Verify that the file was created.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine($"Merged HTML document saved successfully to '{outputPath}'.");
        }
        else
        {
            throw new InvalidOperationException("Failed to save the merged HTML document.");
        }
    }
}
