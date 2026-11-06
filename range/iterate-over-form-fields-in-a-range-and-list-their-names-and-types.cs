using System;
using Aspose.Words;
using Aspose.Words.Fields;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Words.Range;

public class Program
{
    public static void Main()
    {
        // Path for the sample document.
        string filePath = "FormFields.docx";

        // -----------------------------------------------------------------
        // Create a sample document containing different types of form fields.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Form fields:");
        // Text input field.
        builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Default text", 0);
        // Check box field.
        builder.InsertCheckBox("CheckBox1", true, 0);
        // Combo box (drop‑down) field. Use named arguments to guarantee correct overload.
        builder.InsertComboBox(name: "DropDown1", selectedIndex: 0,
            items: new string[] { "Option1", "Option2", "Option3" });

        // Save the document to disk.
        doc.Save(filePath);

        // -----------------------------------------------------------------
        // Load the document back and obtain its whole‑document range.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(filePath);
        AsposeRange range = loadedDoc.Range;

        // -----------------------------------------------------------------
        // Iterate over all form fields that belong to the range and output
        // their names and types.
        // -----------------------------------------------------------------
        foreach (FormField field in range.FormFields)
        {
            Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
        }
    }
}
