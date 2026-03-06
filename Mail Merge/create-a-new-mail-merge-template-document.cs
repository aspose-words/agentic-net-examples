using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample MERGEFIELD fields that will be used during mail merge.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Mark this document as a mail‑merge template (form‑letter type).
        MailMergeSettings settings = doc.MailMergeSettings;
        settings.MainDocumentType = MailMergeMainDocumentType.FormLetters;

        // Save the template to disk.
        doc.Save("MailMergeTemplate.docx");
    }
}
