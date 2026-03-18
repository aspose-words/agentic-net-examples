using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the document that contains the checkbox form field.
        Document doc = new Document("FormFields.docx");

        // Get the checkbox form field by its bookmark/name.
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
