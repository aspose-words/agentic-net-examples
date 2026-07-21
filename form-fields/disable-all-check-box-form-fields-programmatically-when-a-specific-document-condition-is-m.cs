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

        // Add a marker text that will be used as the condition.
        builder.Writeln("Form example. Condition: ENABLE_CHECKBOXES");

        // Insert the first checkbox form field.
        builder.Write("Option 1: ");
        FormField checkBox1 = builder.InsertCheckBox("CheckBox1", false, 20);
        builder.Writeln();

        // Insert the second checkbox form field.
        builder.Write("Option 2: ");
        FormField checkBox2 = builder.InsertCheckBox("CheckBox2", true, 20);
        builder.Writeln();

        // Insert a regular text input field (non‑checkbox) for contrast.
        builder.Write("Enter name: ");
        FormField textInput = builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.Writeln();

        // Save the initial document (optional, demonstrates creation).
        const string initialPath = "FormFields.docx";
        doc.Save(initialPath);

        // Determine whether the specific condition is met.
        // Here the condition is the presence of the marker text "ENABLE_CHECKBOXES".
        bool conditionMet = doc.GetText().Contains("ENABLE_CHECKBOXES");

        if (conditionMet)
        {
            // Iterate through all form fields in the document.
            foreach (FormField field in doc.Range.FormFields)
            {
                // Identify checkbox form fields by their field type.
                if (field.Type == FieldType.FieldFormCheckBox)
                {
                    // Disable the checkbox so it cannot be edited in Word.
                    field.Enabled = false;
                }
            }
        }

        // Save the modified document where all checkboxes are disabled.
        const string outputPath = "FormFields_Disabled.docx";
        doc.Save(outputPath);
    }
}
