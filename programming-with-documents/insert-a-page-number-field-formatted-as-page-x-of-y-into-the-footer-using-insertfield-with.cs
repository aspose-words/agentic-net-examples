using System;
using Aspose.Words;
using Aspose.Words.Fields;   // Required for the FieldType enum

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Insert the page number field in the format "Page X of Y".
        builder.Write("Page ");
        builder.InsertField(FieldType.FieldPage, true);          // PAGE field
        builder.Write(" of ");
        builder.InsertField(FieldType.FieldNumPages, true);     // NUMPAGES field

        // Update all fields so that the page numbers are calculated.
        doc.UpdateFields();

        // Save the document.
        doc.Save("PageNumberFooter.docx");
    }
}
