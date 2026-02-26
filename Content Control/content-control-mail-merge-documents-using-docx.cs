using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control for the "FullName" field.
        StructuredDocumentTag fullNameTag = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        fullNameTag.Title = "FullName";               // Title is used as the mail‑merge field name.
        fullNameTag.PlaceholderName = "FullName";

        // Add a line break between the controls.
        builder.Writeln();

        // Insert a plain‑text content control for the "Address" field.
        StructuredDocumentTag addressTag = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        addressTag.Title = "Address";
        addressTag.PlaceholderName = "Address";

        // Prepare mail‑merge data. Field names must match the titles of the content controls.
        string[] fieldNames = { "FullName", "Address" };
        object[] fieldValues = { "James Bond", "Secret Service Headquarters" };

        // Execute the mail merge; the content controls will be populated with the values.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the resulting document.
        doc.Save("ContentControlMailMerge.docx");
    }
}
