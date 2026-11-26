using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field.
        // Parameters: field name, default checked state, size (points) as an integer.
        builder.InsertCheckBox("AcceptTerms", true, 15);

        // Add descriptive text after the checkbox.
        builder.Writeln(" I accept the terms and conditions.");

        // Save the document to disk.
        doc.Save("CheckboxFormField.docx");
    }
}
