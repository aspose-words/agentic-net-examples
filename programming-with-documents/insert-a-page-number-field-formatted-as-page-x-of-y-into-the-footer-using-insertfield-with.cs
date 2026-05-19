using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsPageNumberFooter
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder cursor to the primary footer of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // Optional: center the footer text.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Write the static part of the page number text.
            builder.Write("Page ");

            // Insert a PAGE field that shows the current page number.
            // The second argument 'true' updates the field result immediately.
            builder.InsertField(FieldType.FieldPage, true);

            // Write the separator text.
            builder.Write(" of ");

            // Insert a NUMPAGES field that shows the total number of pages.
            builder.InsertField(FieldType.FieldNumPages, true);

            // Ensure all fields are up‑to‑date (useful if the document is later modified).
            doc.UpdateFields();

            // Save the document to a file in the same folder as the executable.
            doc.Save("PageNumberFooter.docx");
        }
    }
}
