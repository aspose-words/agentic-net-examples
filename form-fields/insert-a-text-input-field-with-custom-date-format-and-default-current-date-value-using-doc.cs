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

        // Define the name, custom date format, and default value (current date).
        string fieldName = "DateField";
        string dateFormat = "MM/dd/yyyy";
        string defaultValue = DateTime.Now.ToString(dateFormat);

        // Insert a text input form field with the custom date format and default value.
        // Use the overload that requires TextFormFieldType and maxLength.
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, dateFormat, defaultValue, 0);

        // Save the document to disk.
        doc.Save("Output.docx");
    }
}
