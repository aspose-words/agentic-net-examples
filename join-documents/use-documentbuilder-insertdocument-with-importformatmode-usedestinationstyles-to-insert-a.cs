using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System.Drawing;

public class InsertDocumentExample
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string mergedHtmlPath = Path.Combine(outputDir, "Merged.html");

        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        // Define a custom style named "MyStyle" in the template.
        Style templateStyle = templateDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
        templateStyle.Font.Size = 16;
        templateStyle.Font.Name = "Arial";
        templateStyle.Font.Color = Color.Blue;

        // Apply the custom style.
        templateBuilder.ParagraphFormat.StyleName = "MyStyle";
        templateBuilder.Writeln("Template Document Heading");

        // Add a normal paragraph.
        templateBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        templateBuilder.Writeln("This is the body of the template document.");

        // Save the template.
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // ---------- Create the source document ----------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // Define a style with the same name but different formatting.
        Style sourceStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
        sourceStyle.Font.Size = 20;
        sourceStyle.Font.Name = "Times New Roman";
        sourceStyle.Font.Color = Color.Red;

        // Apply the source style.
        sourceBuilder.ParagraphFormat.StyleName = "MyStyle";
        sourceBuilder.Writeln("Source Document Heading");

        // Add a normal paragraph.
        sourceBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        sourceBuilder.Writeln("This is the body of the source document.");

        // Save the source document.
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------- Load the template and insert the source document ----------
        Document mergedDoc = new Document(templatePath);
        DocumentBuilder insertBuilder = new DocumentBuilder(mergedDoc);

        // Move cursor to the end of the template.
        insertBuilder.MoveToDocumentEnd();

        // Optional: insert a page break before the inserted content.
        insertBuilder.InsertBreak(BreakType.PageBreak);

        // Load the source document to be inserted.
        Document docToInsert = new Document(sourcePath);

        // Insert the source document using UseDestinationStyles to adopt the template's styles.
        insertBuilder.InsertDocument(docToInsert, ImportFormatMode.UseDestinationStyles);

        // ---------- Save the merged document as HTML ----------
        mergedDoc.Save(mergedHtmlPath, SaveFormat.Html);

        // Verify that the HTML file was created.
        if (!File.Exists(mergedHtmlPath))
        {
            throw new InvalidOperationException("The merged HTML file was not created.");
        }

        // Inform the user (optional, not required for non‑interactive execution).
        Console.WriteLine("Merged document saved as HTML at:");
        Console.WriteLine(mergedHtmlPath);
    }
}
