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

        // Insert a checkbox form field with a default size (0 = automatic).
        // Name the field so we can retrieve it later.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);
        // Enable explicit size handling.
        checkBox.IsCheckBoxExactSize = true;
        // Set an initial visual size (e.g., 15 points).
        checkBox.CheckBoxSize = 15.0;

        // Retrieve the same checkbox from the form fields collection.
        FormField? retrieved = doc.Range.FormFields["MyCheckBox"];
        if (retrieved == null)
            throw new InvalidOperationException("The expected checkbox form field was not found.");

        // Change the size to improve visual consistency (e.g., 30 points).
        retrieved.CheckBoxSize = 30.0;

        // Save the modified document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ModifiedCheckBox.docx");
        doc.Save(outputPath);
    }
}
