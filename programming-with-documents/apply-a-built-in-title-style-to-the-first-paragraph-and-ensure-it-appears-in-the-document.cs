using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first paragraph with some text.
        builder.Writeln("My Document Title");

        // Apply the built‑in "Title" style to the paragraph.
        builder.ParagraphFormat.StyleName = "Title";

        // Ensure the paragraph appears in the document outline by setting an outline level.
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;

        // Define the output file path (in the current directory).
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TitleStyleOutline.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
