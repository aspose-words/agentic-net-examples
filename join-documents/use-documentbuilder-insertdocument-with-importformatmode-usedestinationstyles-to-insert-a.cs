using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string resultHtmlPath = Path.Combine(outputDir, "Result.html");

        // ---------- Create template document ----------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        // Define a custom style in the template
        Style templateStyle = templateDoc.Styles.Add(StyleType.Paragraph, "MyHeading");
        templateStyle.Font.Name = "Arial";
        templateStyle.Font.Size = 24;
        templateStyle.Font.Color = Color.Blue;

        // Use the custom style
        templateBuilder.ParagraphFormat.StyleName = templateStyle.Name;
        templateBuilder.Writeln("Template Heading");
        templateBuilder.Writeln("Template body text.");

        // Save the template
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // ---------- Create source document ----------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // Define a style with the same name but different formatting
        Style sourceStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "MyHeading");
        sourceStyle.Font.Name = "Times New Roman";
        sourceStyle.Font.Size = 20;
        sourceStyle.Font.Color = Color.Red;

        // Use the source style
        sourceBuilder.ParagraphFormat.StyleName = sourceStyle.Name;
        sourceBuilder.Writeln("Source Heading");
        sourceBuilder.Writeln("Source body text.");

        // Save the source document
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------- Insert source into template ----------
        Document dst = new Document(templatePath);
        Document src = new Document(sourcePath);
        DocumentBuilder builder = new DocumentBuilder(dst);

        // Move cursor to the end and insert a page break (optional)
        builder.MoveToDocumentEnd();
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the source document using destination styles
        builder.InsertDocument(src, ImportFormatMode.UseDestinationStyles);

        // ---------- Save merged document as HTML ----------
        dst.Save(resultHtmlPath, SaveFormat.Html);

        // Verify that the HTML file was created
        if (!File.Exists(resultHtmlPath))
        {
            throw new InvalidOperationException("The HTML output file was not created.");
        }
    }
}
