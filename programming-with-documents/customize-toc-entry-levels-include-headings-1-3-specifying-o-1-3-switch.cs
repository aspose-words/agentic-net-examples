using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC field that includes heading levels 1‑3.
        // \\o "1-3" – specifies the range of heading levels.
        // \\h – makes entries clickable hyperlinks.
        // \\z – hides page numbers in web layout.
        // \\u – builds the TOC using outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add sample headings of various levels.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");

        // This heading is level 4 and should NOT appear in the TOC because of the \\o switch.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;
        builder.Writeln("Heading 4 (excluded)");

        // Update all fields so the TOC reflects the inserted headings.
        doc.UpdateFields();

        // Save the document.
        doc.Save("TOC_Headings1to3.docx");
    }
}
