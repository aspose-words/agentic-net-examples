using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "AsposeJoinDemo");
        Directory.CreateDirectory(workDir);

        // Paths for the template and the document to be inserted.
        string templatePath = Path.Combine(workDir, "Template.docx");
        string sourcePath = Path.Combine(workDir, "Source.docx");
        string outputHtmlPath = Path.Combine(workDir, "Result.html");

        // -----------------------------------------------------------------
        // 1. Create a styled template document.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        // Define a custom paragraph style in the template.
        Style templateStyle = templateDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
        templateStyle.Font.Name = "Arial";
        templateStyle.Font.Size = 16;
        templateStyle.Font.Color = System.Drawing.Color.Blue;

        // Apply the style and write some content.
        templateBuilder.ParagraphFormat.StyleName = templateStyle.Name;
        templateBuilder.Writeln("This is the template document.");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a source document that will be inserted.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // Create a style with the same name but different formatting.
        Style sourceStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
        sourceStyle.Font.Name = "Times New Roman";
        sourceStyle.Font.Size = 14;
        sourceStyle.Font.Color = System.Drawing.Color.Red;

        // Apply the style and write some content.
        sourceBuilder.ParagraphFormat.StyleName = sourceStyle.Name;
        sourceBuilder.Writeln("This is the inserted document.");

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 3. Load the template and insert the source document using
        //    ImportFormatMode.UseDestinationStyles.
        // -----------------------------------------------------------------
        Document resultDoc = new Document(templatePath);
        DocumentBuilder resultBuilder = new DocumentBuilder(resultDoc);

        // Move the cursor to the end of the template document.
        resultBuilder.MoveToDocumentEnd();

        // Insert a page break for visual separation (optional).
        resultBuilder.InsertBreak(BreakType.PageBreak);

        // Load the source document to be inserted.
        Document docToInsert = new Document(sourcePath);

        // Insert the source document, forcing the use of destination styles.
        resultBuilder.InsertDocument(docToInsert, ImportFormatMode.UseDestinationStyles);

        // -----------------------------------------------------------------
        // 4. Save the combined document as HTML.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        resultDoc.Save(outputHtmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Simple validation that the HTML file was created.
        // -----------------------------------------------------------------
        if (File.Exists(outputHtmlPath))
        {
            Console.WriteLine("HTML file successfully created at:");
            Console.WriteLine(outputHtmlPath);
        }
        else
        {
            throw new InvalidOperationException("Failed to create the HTML output file.");
        }
    }
}
