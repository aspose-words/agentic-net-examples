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

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);
        textField.Result = "John Doe";

        // Insert a checkbox form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptTerms", false, 15);
        checkBox.Checked = false;

        // Insert a combo box (dropdown) form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Select a fruit: ");
        FormField comboBox = builder.InsertComboBox("FruitChoice", new[] { "Apple", "Banana", "Cherry" }, 0);
        comboBox.Result = "Apple";

        // Ensure the document contains at least one form field.
        if (doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("The document does not contain any form fields.");
        }

        // Iterate through all form fields and log their Result values.
        foreach (FormField field in doc.Range.FormFields)
        {
            if (field != null)
            {
                // For debugging purposes, output the field name and its current result.
                Console.WriteLine($"Field Name: {field.Name}, Result: {field.Result}");
            }
        }

        // Save the document (even though we only read fields, saving satisfies the lifecycle rule).
        doc.Save("FormFields.docx");
    }
}
