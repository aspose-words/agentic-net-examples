using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add paragraphs with different built‑in styles.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("This is a quote paragraph that should be removed.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("Another quote paragraph to delete.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1 paragraph.");

        // Remove all paragraphs that use the Quote style.
        Paragraph[] paragraphs = doc.FirstSection.Body.Paragraphs.ToArray();
        foreach (Paragraph para in paragraphs)
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Quote)
                para.Remove();
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
