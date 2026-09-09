using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.html");

        // -----------------------------------------------------------------
        // 1. Create a styled template document.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        // Apply a heading style to the first paragraph.
        templateBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        templateBuilder.Writeln("Template Title");

        // Add a normal paragraph.
        templateBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        templateBuilder.Writeln("This is the template content.");

        // Save the template so it can be loaded later.
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a source document that will be inserted.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // Use the same style name but different formatting to demonstrate style clash handling.
        Style customStyle = sourceBuilder.Document.Styles.Add(StyleType.Paragraph, "CustomStyle");
        customStyle.Font.Name = "Courier New";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = System.Drawing.Color.DarkRed;

        sourceBuilder.ParagraphFormat.StyleName = customStyle.Name;
        sourceBuilder.Writeln("Source document paragraph with custom style.");

        // Save the source document.
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load both documents and insert the source into the template.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        Document loadedSource = new Document(sourcePath);

        DocumentBuilder builder = new DocumentBuilder(loadedTemplate);
        builder.MoveToDocumentEnd();
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the source document using UseDestinationStyles to adopt the template's styles.
        builder.InsertDocument(loadedSource, ImportFormatMode.UseDestinationStyles);

        // -----------------------------------------------------------------
        // 4. Save the merged result as HTML.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPath, SaveFormat.Html);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The merged HTML file was not created.");
        }

        // Optional: Inform that the process completed successfully.
        Console.WriteLine("Document merged and saved as HTML at: " + outputPath);
    }
}
