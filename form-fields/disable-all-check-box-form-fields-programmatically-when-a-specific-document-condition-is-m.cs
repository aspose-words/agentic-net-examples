using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three checkbox form fields with distinct names.
        builder.InsertCheckBox("CheckBox1", true, 0);
        builder.Writeln(); // Add line break for readability.
        builder.InsertCheckBox("CheckBox2", false, 0);
        builder.Writeln();
        builder.InsertCheckBox("CheckBox3", false, 0);
        builder.Writeln();

        // Save the initial document (optional, for inspection).
        doc.Save("Initial.docx");

        // Access the collection of form fields.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields == null || formFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were found in the document.");
        }

        // Determine if the specific condition is met.
        // Condition: any checkbox is checked.
        bool anyChecked = false;
        foreach (FormField field in formFields)
        {
            if (field.Type == FieldType.FieldFormCheckBox && field.Checked)
            {
                anyChecked = true;
                break;
            }
        }

        // If the condition is met, disable all checkbox form fields.
        if (anyChecked)
        {
            foreach (FormField field in formFields)
            {
                if (field.Type == FieldType.FieldFormCheckBox)
                {
                    field.Enabled = false;
                }
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");

        // Output status to console (no user interaction required).
        Console.WriteLine("Document processing completed. Checkboxes disabled: " + anyChecked);
    }
}
