using System;
using System.Text;
using System.Collections;
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
        FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
        comboBox.CalculateOnExit = true;

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Click this check box to tick/untick it: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
        checkBox.IsCheckBoxExactSize = true;
        checkBox.HelpText = "Right click to check this box";
        checkBox.OwnHelp = true;
        checkBox.StatusText = "Checkbox status text";
        checkBox.OwnStatus = true;

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter text here: ");
        FormField textInput = builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder text", 50);
        textInput.EntryMacro = "EntryMacro";
        textInput.ExitMacro = "ExitMacro";
        textInput.TextInputDefault = "Regular";
        textInput.TextInputFormat = "FIRST CAPITAL";
        textInput.SetTextInputValue("New placeholder text");

        // Use a visitor to collect information about all form fields.
        FormFieldVisitor visitor = new FormFieldVisitor();
        FormFieldCollection formFields = doc.Range.FormFields;
        using (IEnumerator<FormField> enumerator = formFields.GetEnumerator())
        {
            while (enumerator.MoveNext())
                enumerator.Current.Accept(visitor);
        }

        // Print the collected information to the console.
        Console.WriteLine(visitor.GetText());

        // Update fields and save the document as HTML.
        doc.UpdateFields();
        doc.Save("FormFields.html");
    }
}

// Visitor that builds a plain‑text description of each form field.
class FormFieldVisitor : DocumentVisitor
{
    private readonly StringBuilder _builder = new StringBuilder();

    public override VisitorAction VisitFormField(FormField formField)
    {
        _builder.AppendLine($"{formField.Type}: \"{formField.Name}\"");
        _builder.AppendLine($"\tStatus: {(formField.Enabled ? "Enabled" : "Disabled")}");
        _builder.AppendLine($"\tHelp Text: {formField.HelpText}");
        _builder.AppendLine($"\tEntry macro name: {formField.EntryMacro}");
        _builder.AppendLine($"\tExit macro name: {formField.ExitMacro}");

        switch (formField.Type)
        {
            case FieldType.FieldFormDropDown:
                _builder.AppendLine($"\tDrop-down items count: {formField.DropDownItems.Count}, default selected index: {formField.DropDownSelectedIndex}");
                _builder.AppendLine($"\tDrop-down items: {string.Join(", ", formField.DropDownItems)}");
                break;
            case FieldType.FieldFormCheckBox:
                _builder.AppendLine($"\tCheckbox size: {formField.CheckBoxSize}");
                _builder.AppendLine($"\tCheckbox is currently: {(formField.Checked ? "checked" : "unchecked")}, by default: {(formField.Default ? "checked" : "unchecked")}");
                break;
            case FieldType.FieldFormTextInput:
                _builder.AppendLine($"\tInput format: {formField.TextInputFormat}");
                _builder.AppendLine($"\tCurrent contents: {formField.Result}");
                break;
        }

        return VisitorAction.Continue;
    }

    public string GetText()
    {
        return _builder.ToString();
    }
}
