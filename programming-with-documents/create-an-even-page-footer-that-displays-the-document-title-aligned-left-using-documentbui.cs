using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a title for the document – this will be displayed by the TITLE field.
        doc.BuiltInDocumentProperties.Title = "My Document Title";

        // Enable different footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Move the builder cursor to the even-page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

        // Align the paragraph to the left.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // Insert a TITLE field that displays the document title.
        // The field is updated immediately (second argument = true).
        builder.InsertField(FieldType.FieldTitle, true);

        // Return to the main body of the document.
        builder.MoveToSection(0);

        // Add a few pages to demonstrate the even-page footer.
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Save the document to the local file system.
        doc.Save("EvenFooter.docx");
    }
}
