using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field with a known name.
        // The checkbox is initially unchecked (false) and uses the default size (0).
        builder.Write("Please check the box: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Save the document so that the form field is persisted.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CheckboxFormField.docx");
        doc.Save(outputPath);

        // Reload the document to simulate a separate read operation.
        Document loadedDoc = new Document(outputPath);

        // Access the form field collection and locate the checkbox by its name.
        FormField loadedCheckBox = loadedDoc.Range.FormFields["MyCheckBox"];
        if (loadedCheckBox == null)
        {
            throw new InvalidOperationException("The expected checkbox form field was not found.");
        }

        // The Result property of a checkbox contains "1" if checked, otherwise "0".
        string result = loadedCheckBox.Result;
        bool isChecked = result == "1";

        // Output the determination to the console.
        Console.WriteLine($"Checkbox '{loadedCheckBox.Name}' is {(isChecked ? "checked" : "unchecked")} (Result = \"{result}\").");
    }
}
