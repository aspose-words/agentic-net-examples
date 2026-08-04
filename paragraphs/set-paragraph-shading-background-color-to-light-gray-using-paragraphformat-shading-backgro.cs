using System;
using System.IO;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph shading background color to light gray.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Write some text into the paragraph.
        builder.Writeln("This paragraph has a light gray background shading.");

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "ParagraphShading.docx");
        doc.Save(outputPath);
    }
}
