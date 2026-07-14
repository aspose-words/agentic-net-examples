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

        // Insert a checkbox form field with an initial size.
        builder.Write("Original checkbox: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 20);
        // Enable explicit size handling.
        checkBox.IsCheckBoxExactSize = true;
        builder.InsertParagraph();

        // Retrieve the same checkbox from the form fields collection.
        FormField? retrieved = doc.Range.FormFields["MyCheckBox"];
        if (retrieved == null)
            throw new InvalidOperationException("Checkbox form field not found.");

        // Change the size of the checkbox to improve visual consistency.
        retrieved.IsCheckBoxExactSize = true; // Ensure exact size is enabled.
        retrieved.CheckBoxSize = 40; // New size in points.

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CheckboxSizeChanged.docx");
        doc.Save(outputPath);
    }
}
