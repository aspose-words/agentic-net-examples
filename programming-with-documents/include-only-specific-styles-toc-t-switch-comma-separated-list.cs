using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class TocWithCustomStyles
{
    public static void Main(string[] args)
    {
        CreateDocument();
    }

    public static void CreateDocument()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC field. We will configure it after insertion.
        FieldToc toc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);

        // Include only the custom styles "Quote" and "Intense Quote" in the TOC.
        // The styles are separated by a comma, which is the default separator.
        toc.CustomStyles = "Quote, Intense Quote";

        // Optionally, limit the heading levels that the TOC will consider.
        toc.HeadingLevelRange = "1-3";

        // Insert some paragraphs with built‑in heading styles – they will be ignored because we only
        // specified custom styles in the \t switch.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Built‑in Heading 1 (will not appear)");

        // Insert a paragraph with the "Quote" style – this will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("Quote style entry");

        // Insert a paragraph with the "Intense Quote" style – this will also appear.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.IntenseQuote;
        builder.Writeln("Intense Quote style entry");

        // Insert a paragraph with a different custom style – it will not be listed.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Normal style entry (will not appear)");

        // Update the TOC field so that it reflects the current document content.
        toc.UpdatePageNumbers();
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save("TocWithCustomStyles.docx");
    }
}
