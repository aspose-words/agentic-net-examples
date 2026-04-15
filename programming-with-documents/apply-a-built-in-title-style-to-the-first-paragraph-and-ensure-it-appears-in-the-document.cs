using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in "Title" style to the first paragraph.
        builder.ParagraphFormat.StyleName = "Title";

        // Set the outline level so the paragraph appears in the document outline (e.g., level 1).
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;

        // Write the title text.
        builder.Writeln("My Document Title");

        // Reset formatting for subsequent paragraphs if needed.
        builder.ParagraphFormat.ClearFormatting();

        // Add a normal paragraph for demonstration.
        builder.Writeln("This is the body of the document.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TitleStyleOutline.docx");
        doc.Save(outputPath);
    }
}
