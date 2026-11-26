using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample form fields.
        builder.Writeln("Please fill the form below:");
        builder.InsertField("FORMTEXT Name", "John Doe");
        builder.Writeln();
        builder.InsertField("FORMCHECKBOX Agree", "0");

        // Save the original document (optional, for verification).
        doc.Save("Original.docx");

        // Delete all form fields by iterating over the document's range.
        FormFieldCollection formFields = doc.Range.FormFields;
        while (formFields.Count > 0)
        {
            // Remove the first form field; the collection updates automatically.
            formFields[0].Remove();
        }

        // Save the document after removal of form fields.
        doc.Save("Cleaned.docx");
    }
}
