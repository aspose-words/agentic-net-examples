using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample long text to demonstrate justification and word wrap.
        string sampleText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.";

        // Set paragraph formatting: justified alignment and enable word wrap.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        builder.ParagraphFormat.WordWrap = true; // Explicitly enable word wrap (default is true).

        // Write the text into the paragraph.
        builder.Writeln(sampleText);

        // Determine output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "JustifiedParagraph.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
