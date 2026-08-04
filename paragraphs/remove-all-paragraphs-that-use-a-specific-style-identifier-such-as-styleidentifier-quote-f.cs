using System;
using Aspose.Words;

public class RemoveQuoteParagraphs
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph.");

        // Add a paragraph with the built‑in Quote style (the target for removal).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("This is a quote paragraph that should be removed.");

        // Add another normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Another normal paragraph.");

        // Add a heading paragraph (should remain untouched).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        // Remove all paragraphs that use the Quote style.
        // Iterate over a copy of the paragraph collection to safely modify the document.
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs.ToArray())
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Quote)
                para.Remove();
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
