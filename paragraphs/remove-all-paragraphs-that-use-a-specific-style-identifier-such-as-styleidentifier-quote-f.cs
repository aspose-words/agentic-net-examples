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
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("This is a quote paragraph that should be removed.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph that should stay.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("Another quote paragraph to be removed.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading paragraph that should stay.");

        // Remove all paragraphs that use the Quote style.
        // Iterate backwards to safely remove nodes while traversing.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (int i = paragraphs.Count - 1; i >= 0; i--)
        {
            Paragraph para = (Paragraph)paragraphs[i];
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Quote)
            {
                para.Remove();
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
