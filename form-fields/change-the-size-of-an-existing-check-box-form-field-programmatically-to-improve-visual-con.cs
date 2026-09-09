using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field with the default automatic size.
        // The name "MyCheckBox" will be used later to locate the field.
        builder.Write("Please tick the box: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Ensure the checkbox size can be set explicitly.
        checkBox.IsCheckBoxExactSize = true;

        // Change the size of the existing checkbox to 30 points.
        // This improves visual consistency with other form elements.
        checkBox.CheckBoxSize = 30.0;

        // Optional: verify that the size was applied.
        if (Math.Abs(checkBox.CheckBoxSize - 30.0) > 0.001)
            throw new InvalidOperationException("Failed to set the checkbox size.");

        // Save the document to disk.
        doc.Save("CheckboxSizeChanged.docx");
    }
}
