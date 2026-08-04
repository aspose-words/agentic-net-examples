using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some regular text.
        builder.Writeln("Sample document.");

        // Marker that determines whether check boxes should be disabled.
        builder.Writeln("Condition: DisableAll");

        // Add a paragraph before the form fields.
        builder.Writeln("Form fields below:");

        // Insert first check box.
        builder.Write("CheckBox 1: ");
        FormField checkBox1 = builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();

        // Insert second check box.
        builder.Write("CheckBox 2: ");
        FormField checkBox2 = builder.InsertCheckBox("CheckBox2", true, 0);
        builder.Writeln();

        // Determine if the specific condition is present in the document.
        bool conditionMet = doc.GetText().Contains("DisableAll");

        if (conditionMet)
        {
            // Access the collection of all form fields.
            FormFieldCollection formFields = doc.Range.FormFields;

            // Validate that at least one form field exists.
            if (formFields.Count == 0)
                throw new InvalidOperationException("The document does not contain any form fields.");

            // Disable every check box form field.
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
    }
}
