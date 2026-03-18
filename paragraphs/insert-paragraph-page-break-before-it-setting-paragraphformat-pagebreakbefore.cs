using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the flag that forces a page break before the next paragraph.
        builder.ParagraphFormat.PageBreakBefore = true;

        // Insert the paragraph that will appear on a new page.
        builder.Writeln("This paragraph starts on a new page.");

        // Save the resulting document.
        doc.Save("ParagraphWithPageBreakBefore.docx");
    }
}
