using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add paragraphs with different built‑in styles.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("This paragraph uses the Quote style and will be removed.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Another normal paragraph.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("Second Quote style paragraph to be removed.");

        // Remove all paragraphs that have the Quote style.
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs.ToArray())
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Quote)
                para.Remove();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
