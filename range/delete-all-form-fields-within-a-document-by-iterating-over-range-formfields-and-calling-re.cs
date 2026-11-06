using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a couple of form fields for demonstration.
        builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Default1", 0);
        builder.Writeln(); // Add a line break.
        builder.InsertCheckBox("CheckBox1", true, 0);

        // Save the original document (optional, shows before removal).
        doc.Save("Original.docx");

        // Delete all form fields by iterating over the FormFields collection.
        // Iterate backwards because removing a field changes the collection indices.
        for (int i = doc.Range.FormFields.Count - 1; i >= 0; i--)
        {
            doc.Range.FormFields[i].Remove();
        }

        // Save the document after all form fields have been removed.
        doc.Save("NoFormFields.docx");
    }
}
