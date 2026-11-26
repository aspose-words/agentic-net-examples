using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several legacy form fields.
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();

        // Insert a text input form field using the correct enum.
        builder.InsertTextInput(
            "TextInput1",
            TextFormFieldType.Regular,
            "",
            "Default text",
            0);
        builder.Writeln();

        builder.InsertComboBox("ComboBox1", new string[] { "OptionA", "OptionB", "OptionC" }, 0);
        builder.Writeln();

        // Retrieve the count of form fields within the document's range.
        int formFieldCount = doc.Range.FormFields.Count;

        // Output the result.
        Console.WriteLine($"Number of form fields in the document range: {formFieldCount}");
    }
}
