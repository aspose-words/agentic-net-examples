using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Optional: center the footer text.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Insert a formula field that evaluates to the current page number.
        // The field code " = PAGE " tells Word to calculate the PAGE value.
        builder.InsertField("= PAGE", "");

        // Save the document.
        doc.Save("FooterPageNumber.docx");
    }
}
