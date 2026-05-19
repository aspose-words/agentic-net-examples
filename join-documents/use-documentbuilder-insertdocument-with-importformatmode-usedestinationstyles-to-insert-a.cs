using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string templatePath = Path.Combine(outputDir, "Template.docx");
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string resultPath = Path.Combine(outputDir, "Result.html");

        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        // Define a custom style named "MyStyle" with specific formatting.
        Style templateStyle = templateDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
        templateStyle.Font.Size = 16;
        templateStyle.Font.Color = System.Drawing.Color.Blue;
        templateStyle.Font.Name = "Arial";

        // Apply the style and write some content.
        templateBuilder.ParagraphFormat.StyleName = templateStyle.Name;
        templateBuilder.Writeln("Template heading using MyStyle.");

        // Save the template as DOCX.
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // ---------- Create the source document ----------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // Define a style with the same name but different formatting.
        Style sourceStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
        sourceStyle.Font.Size = 12;
        sourceStyle.Font.Color = System.Drawing.Color.Red;
        sourceStyle.Font.Name = "Times New Roman";

        // Apply the style and write some content.
        sourceBuilder.ParagraphFormat.StyleName = sourceStyle.Name;
        sourceBuilder.Writeln("Source paragraph using MyStyle.");

        // Save the source as DOCX.
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------- Load documents and insert source into template ----------
        Document template = new Document(templatePath);
        Document source = new Document(sourcePath);

        DocumentBuilder builder = new DocumentBuilder(template);
        builder.MoveToDocumentEnd();

        // Insert the source document using destination styles.
        builder.InsertDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as HTML.
        template.Save(resultPath, SaveFormat.Html);

        // Validate that the HTML file was created.
        if (!File.Exists(resultPath))
        {
            throw new Exception("Failed to create the HTML output file.");
        }

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine("Document merged and saved as HTML successfully.");
    }
}
