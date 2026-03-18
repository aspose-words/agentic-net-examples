using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        // Center the footer text.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Insert a DATE field that will display the current date.
        builder.InsertField(FieldType.FieldDate, true);
        builder.Write(" - Page ");

        // Insert a PAGE field that will display the current page number.
        builder.InsertField(FieldType.FieldPage, true);

        // Ensure all fields have up‑to‑date results before saving.
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save("DynamicFooter.docx");
    }
}
