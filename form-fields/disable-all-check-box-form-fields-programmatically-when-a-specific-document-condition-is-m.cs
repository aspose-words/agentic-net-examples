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

        // Add a marker text that will be used as the condition to disable check boxes.
        builder.Writeln("Document condition: DisableCheckBoxes");

        // Insert some sample check box form fields.
        builder.Write("Option 1: ");
        FormField checkBox1 = builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();

        builder.Write("Option 2: ");
        FormField checkBox2 = builder.InsertCheckBox("CheckBox2", true, 0);
        builder.Writeln();

        // Insert a non‑checkbox form field to demonstrate that only check boxes are affected.
        builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Enter text", 50);
        builder.Writeln();

        // Determine whether the specific condition is present in the document.
        // Here we simply look for the marker text added earlier.
        bool conditionMet = doc.GetText().Contains("DisableCheckBoxes");

        if (conditionMet)
        {
            // Iterate through all form fields in the document.
            FormFieldCollection formFields = doc.Range.FormFields;
            foreach (FormField field in formFields)
            {
                // Check if the field is a check box.
                if (field.Type == FieldType.FieldFormCheckBox)
                {
                    // Disable the check box so it cannot be edited in Word.
                    field.Enabled = false;
                }
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
