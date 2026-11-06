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

        // Insert a text input form field.
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default text", 0);
        builder.Writeln();

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBoxField", false, 0);
        builder.Writeln();

        // Verify that form fields were added.
        FormFieldCollection fields = doc.Range.FormFields;
        if (fields == null || fields.Count == 0)
            throw new InvalidOperationException("No form fields were created in the document.");

        // Protect the document so that only form fields can be edited.
        // ProtectionType enum is available directly in the Aspose.Words namespace.
        doc.Protect(ProtectionType.AllowOnlyFormFields, "password");

        // Save the protected document.
        doc.Save("ProtectedFormFields.docx");
    }
}
