using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a combo box form field.
        builder.Write("Choose a value from this combo box: ");
        FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
        comboBox.CalculateOnExit = true;

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Click this check box to tick/untick it: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
        checkBox.IsCheckBoxExactSize = true;

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter text here: ");
        FormField textInput = builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder text", 50);
        textInput.EntryMacro = "EntryMacro";
        textInput.ExitMacro = "ExitMacro";

        // Get the collection of all form fields in the document's range.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Iterate over the collection and print each field's name and type.
        using (IEnumerator<FormField> enumerator = formFields.GetEnumerator())
        {
            while (enumerator.MoveNext())
            {
                FormField field = enumerator.Current;
                Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
            }
        }
    }
}
