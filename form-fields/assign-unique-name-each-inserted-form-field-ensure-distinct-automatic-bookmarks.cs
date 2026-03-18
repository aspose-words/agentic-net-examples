using System;
using Aspose.Words;
using Aspose.Words.Fields;

class FormFieldUniqueNames
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Counter used to generate distinct names for each form field.
        int fieldIndex = 1;

        // ---------- ComboBox ----------
        builder.Write("Select a fruit: ");
        // Insert a combo box; the name is supplied directly.
        FormField comboBox = builder.InsertComboBox($"ComboBox_{fieldIndex++}", new[] { "Apple", "Banana", "Cherry" }, 0);
        // Explicitly set the Name property to guarantee uniqueness (optional, but clear).
        comboBox.Name = $"ComboBox_{fieldIndex - 1}";

        builder.InsertBreak(BreakType.ParagraphBreak);

        // ---------- CheckBox ----------
        builder.Write("Accept terms: ");
        // Insert a check box with a unique name.
        FormField checkBox = builder.InsertCheckBox($"CheckBox_{fieldIndex++}", false, 50);
        checkBox.Name = $"CheckBox_{fieldIndex - 1}";

        builder.InsertBreak(BreakType.ParagraphBreak);

        // ---------- Text Input ----------
        builder.Write("Enter your name: ");
        // Insert a text input field; the name also creates a bookmark with the same name.
        FormField textInput = builder.InsertTextInput($"TextInput_{fieldIndex++}", TextFormFieldType.Regular, "", "John Doe", 0);
        textInput.Name = $"TextInput_{fieldIndex - 1}";

        // Save the document to disk.
        doc.Save("FormFields_UniqueNames.docx");
    }
}
