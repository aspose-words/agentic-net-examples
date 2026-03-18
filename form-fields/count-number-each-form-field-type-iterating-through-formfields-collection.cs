using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a combo box (dropdown) form field.
        builder.Write("Choose a fruit: ");
        builder.InsertComboBox("FruitCombo", new[] { "Apple", "Banana", "Cherry" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Accept terms: ");
        builder.InsertCheckBox("AcceptCheck", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter name: ");
        builder.InsertTextInput("NameInput", TextFormFieldType.Regular, "", "Your name", 50);

        // Retrieve the collection of all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Dictionary to store the count of each form field type.
        var typeCounts = new Dictionary<FieldType, int>();

        // Iterate through the collection and tally the types.
        foreach (FormField field in formFields)
        {
            FieldType type = field.Type;
            if (typeCounts.ContainsKey(type))
                typeCounts[type]++;
            else
                typeCounts[type] = 1;
        }

        // Output the counts to the console.
        foreach (KeyValuePair<FieldType, int> entry in typeCounts)
        {
            Console.WriteLine($"{entry.Key}: {entry.Value}");
        }

        // Save the document (optional, demonstrates the required save lifecycle step).
        doc.Save("FormFieldsCount.docx");
    }
}
