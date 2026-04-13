using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for inserting content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that will act as the condition trigger.
        builder.Writeln("This document contains a condition to Disable check boxes.");

        // Insert several checkbox form fields with distinct names.
        builder.Write("Option 1: ");
        FormField checkBox1 = builder.InsertCheckBox("CheckBox1", false, 20);
        builder.Writeln();

        builder.Write("Option 2: ");
        FormField checkBox2 = builder.InsertCheckBox("CheckBox2", true, 20);
        builder.Writeln();

        builder.Write("Option 3: ");
        FormField checkBox3 = builder.InsertCheckBox("CheckBox3", false, 20);
        builder.Writeln();

        // Ensure that at least one form field exists before attempting to modify them.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
        {
            throw new InvalidOperationException("The document does not contain any form fields.");
        }

        // Define the specific condition: if the document contains the word "Disable".
        bool conditionMet = doc.GetText().Contains("Disable", StringComparison.OrdinalIgnoreCase);

        // If the condition is met, disable all checkbox form fields.
        if (conditionMet)
        {
            foreach (FormField field in formFields)
            {
                // Validate that the field is a checkbox before disabling.
                if (field.Type == FieldType.FieldFormCheckBox)
                {
                    // Disable the checkbox so it cannot be edited.
                    field.Enabled = false;

                    // Optionally, also uncheck it.
                    field.Checked = false;
                }
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
