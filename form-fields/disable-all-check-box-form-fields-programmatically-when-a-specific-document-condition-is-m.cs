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

        // Insert a text input form field that will act as the condition trigger.
        // The field is named "ConditionField".
        builder.Write("Condition (type 'Disable' to deactivate check boxes): ");
        FormField conditionField = builder.InsertTextInput(
            "ConditionField",
            TextFormFieldType.Regular,
            "",
            "Enable", // placeholder text
            0);
        // Simulate user input by setting the result directly.
        conditionField.Result = "Disable";

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert several check box form fields.
        builder.Write("Option A: ");
        builder.InsertCheckBox("CheckBoxA", false, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        builder.Write("Option B: ");
        builder.InsertCheckBox("CheckBoxB", true, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        builder.Write("Option C: ");
        builder.InsertCheckBox("CheckBoxC", false, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Determine whether the condition is met.
        // The condition is satisfied when the text input field's result equals "Disable".
        bool shouldDisable = false;
        FormField fetchedCondition = doc.Range.FormFields["ConditionField"];
        if (fetchedCondition != null && string.Equals(fetchedCondition.Result, "Disable", StringComparison.OrdinalIgnoreCase))
        {
            shouldDisable = true;
        }

        // If the condition is met, disable all check box form fields.
        if (shouldDisable)
        {
            foreach (FormField field in doc.Range.FormFields)
            {
                // Only process check box fields.
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
