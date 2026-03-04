using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertFormFieldsIntoDotx
{
    static void Main()
    {
        // Load the DOTX template that will receive the form‑field documents.
        Document template = new Document("Template.dotx");

        // Create a DocumentBuilder for the template.
        DocumentBuilder builder = new DocumentBuilder(template);

        // Position the cursor at the end of the template where the inserts will occur.
        builder.MoveToDocumentEnd();

        // Load the first source document that contains form fields.
        Document sourceDoc1 = new Document("FormFields1.docx");

        // Insert the first source document, preserving its original formatting.
        builder.InsertDocument(sourceDoc1, ImportFormatMode.KeepSourceFormatting);

        // Load the second source document that also contains form fields.
        Document sourceDoc2 = new Document("FormFields2.docx");

        // Insert the second source document in the same way.
        builder.InsertDocument(sourceDoc2, ImportFormatMode.KeepSourceFormatting);

        // Update all fields (including the inserted form fields) so their results are current.
        template.UpdateFields();

        // Save the combined document. The output format can be any supported type (e.g., DOCX).
        template.Save("CombinedResult.docx", SaveFormat.Docx);
    }
}
