using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the built‑in Title property – the TITLE field will display this value.
        doc.BuiltInDocumentProperties.Title = "Sample Document Title";

        // Initialize a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Move the builder's cursor to the even‑page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

        // Align the paragraph in the footer to the left.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // Insert a TITLE field (do not update automatically) and then update it
        // so the field result shows the document title.
        FieldTitle titleField = (FieldTitle)builder.InsertField(FieldType.FieldTitle, false);
        titleField.Update();

        // Add some body content with page breaks to demonstrate the even footer.
        builder.MoveToSection(0);
        builder.Writeln("First page (odd).");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page (even).");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page (odd).");

        // Save the document to a file in the current directory.
        doc.Save("EvenFooter.docx");
    }
}
