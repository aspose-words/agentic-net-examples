using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set typical hyperlink character formatting.
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;
        // Apply the built‑in Hyperlink character style.
        builder.Font.StyleIdentifier = StyleIdentifier.Hyperlink;

        // Insert the hyperlink run.
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);

        // Reset font formatting for subsequent text.
        builder.Font.ClearFormatting();

        // End the paragraph.
        builder.Writeln();

        // Save the document.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "HyperlinkParagraph.docx");
        doc.Save(outputPath);
    }
}
