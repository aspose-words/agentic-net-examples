using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will be styled as a quote.
        builder.Writeln("The only limit to our realization of tomorrow is our doubts of today.");

        // Get the paragraph that was just added.
        Paragraph paragraph = builder.CurrentParagraph;

        // Apply the built‑in "Quote" style.
        paragraph.ParagraphFormat.StyleName = "Quote";

        // Increase the left indent for emphasis (value in points).
        paragraph.ParagraphFormat.LeftIndent = 30; // 30 points ≈ 0.42 inches

        // Save the document to a file.
        string outputPath = "QuoteStyled.docx";
        doc.Save(outputPath);
    }
}
