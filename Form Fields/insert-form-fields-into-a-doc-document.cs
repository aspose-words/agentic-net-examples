using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a combo box form field.
        builder.Write("Choose a value from this combo box: ");
        FormField comboBox = builder.InsertComboBox(
            "MyComboBox",                     // field name
            new[] { "One", "Two", "Three" }, // list items
            0);                               // default selected index
        comboBox.CalculateOnExit = true;      // recalculate when the field loses focus

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Click this check box to tick/untick it: ");
        FormField checkBox = builder.InsertCheckBox(
            "MyCheckBox", // field name
            false,        // initial state (unchecked)
            50);          // size in points
        checkBox.IsCheckBoxExactSize = true;
        checkBox.HelpText = "Right click to check this box";
        checkBox.OwnHelp = true;
        checkBox.StatusText = "Checkbox status text";
        checkBox.OwnStatus = true;

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter text here: ");
        FormField textInput = builder.InsertTextInput(
            "MyTextInput",               // field name
            TextFormFieldType.Regular,  // type of text field
            "",                         // default text (empty)
            "Placeholder text",         // placeholder text shown to the user
            50);                        // maximum length
        textInput.EntryMacro = "EntryMacro";
        textInput.ExitMacro = "ExitMacro";
        textInput.TextInputDefault = "Regular";
        textInput.TextInputFormat = "FIRST CAPITAL";
        textInput.SetTextInputValue("New placeholder text");

        // Save the document to disk.
        doc.Save("FormFields.docx");
    }
}
