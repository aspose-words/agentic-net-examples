using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a normal paragraph.
        builder.Writeln("This is a normal paragraph.");

        // Add a paragraph with the built‑in Quote style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("This paragraph uses the Quote style and will be removed.");

        // Add another normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Another normal paragraph.");

        // Remove all paragraphs that use the Quote style.
        // Iterate over a copy of the paragraph collection to avoid modifying the collection while iterating.
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs.ToArray())
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Quote)
                para.Remove();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
