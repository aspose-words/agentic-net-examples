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

        // Add a paragraph that will act as the condition trigger.
        builder.Writeln("Condition: DisableAll");

        // Insert a few checkbox form fields.
        builder.Write("Check 1: ");
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();

        builder.Write("Check 2: ");
        builder.InsertCheckBox("CheckBox2", true, 0);
        builder.Writeln();

        // Determine whether the specific condition is present in the document.
        bool conditionMet = doc.GetText().Contains("DisableAll");

        if (conditionMet)
        {
            // Iterate through all form fields in the document.
            FormFieldCollection formFields = doc.Range.FormFields;
            foreach (FormField field in formFields)
            {
                // Identify checkbox form fields and disable them.
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
