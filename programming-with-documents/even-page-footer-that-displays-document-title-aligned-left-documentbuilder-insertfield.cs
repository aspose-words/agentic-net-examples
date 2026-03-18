using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document and associate a builder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the built‑in Title property – the TITLE field will read this value.
        doc.BuiltInDocumentProperties.Title = "My Document Title";

        // Enable separate footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Move the builder's cursor to the even‑page footer of the current section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

        // Ensure the paragraph is left‑aligned.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // Insert a TITLE field that displays the document title.
        // The second argument (false) tells Aspose.Words not to update the field immediately.
        builder.InsertField(FieldType.FieldTitle, false);

        // Return to the main body of the document.
        builder.MoveToSection(0);

        // Add some sample content to generate multiple pages.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Page {i}");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the resulting document.
        doc.Save("EvenFooterTitle.docx");
    }
}
