using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few form fields into the document.
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();
        builder.InsertComboBox("ComboBox1", new string[] { "Option1", "Option2" }, 0);
        builder.Writeln();
        builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Default text", 0);
        builder.Writeln();

        // Save the document (optional, demonstrates lifecycle handling).
        string filePath = "FormFieldsDocument.docx";
        doc.Save(filePath);

        // Retrieve the range of the whole document.
        Aspose.Words.Range range = doc.Range;

        // Count the form fields within this range.
        int formFieldCount = range.FormFields.Count;

        // Output the count.
        Console.WriteLine($"Number of form fields in the document range: {formFieldCount}");
    }
}
