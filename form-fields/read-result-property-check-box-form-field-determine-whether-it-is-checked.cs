using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new document and insert a checkbox form field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertCheckBox("MyCheckBox", true, 0); // true = checked

        // Retrieve the checkbox form field by its name.
        FormField checkBox = doc.Range.FormFields["MyCheckBox"];
        if (checkBox == null)
        {
            Console.WriteLine("Checkbox form field not found.");
            return;
        }

        // For a checkbox the Result property is "1" when checked and "0" when unchecked.
        string result = checkBox.Result;
        bool isChecked = result == "1";

        Console.WriteLine($"Checkbox \"{checkBox.Name}\" is checked: {isChecked}");
    }
}
