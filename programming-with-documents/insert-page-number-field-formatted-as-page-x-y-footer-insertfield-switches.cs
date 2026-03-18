using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Center‑align the paragraph that will contain the page‑number field.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Insert the literal text "Page ".
        builder.Write("Page ");

        // Insert a PAGE field that shows the current page number.
        // The second argument (true) tells Aspose.Words to update the field immediately.
        builder.InsertField(FieldType.FieldPage, true);

        // Insert the literal text " of ".
        builder.Write(" of ");

        // Insert a NUMPAGES field that shows the total number of pages in the document.
        builder.InsertField(FieldType.FieldNumPages, true);

        // Re‑calculate all fields so that the displayed values are correct.
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save("PageNumberFooter.docx");
    }
}
