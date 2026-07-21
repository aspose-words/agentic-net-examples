using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add paragraphs with various built‑in styles.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a Normal paragraph.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("This is a Quote paragraph that should be removed.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("Another Quote paragraph to be removed.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("This is a Heading 1 paragraph.");

        // Remove all paragraphs that use the Quote style.
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs.ToArray())
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Quote)
                para.Remove();
        }

        // Save the resulting document.
        doc.Save("RemovedQuoteParagraphs.docx");
    }
}
