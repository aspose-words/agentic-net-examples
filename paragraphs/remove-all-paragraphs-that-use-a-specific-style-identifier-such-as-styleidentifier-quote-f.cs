using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add normal paragraph.
        builder.Writeln("First paragraph - normal.");

        // Add paragraphs with the Quote style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("Quote paragraph 1.");
        builder.Writeln("Quote paragraph 2.");

        // Add another normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Another normal paragraph.");

        // Remove all paragraphs that use the Quote style.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (int i = paragraphs.Count - 1; i >= 0; i--)
        {
            Paragraph para = (Paragraph)paragraphs[i];
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Quote)
            {
                para.Remove();
            }
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
