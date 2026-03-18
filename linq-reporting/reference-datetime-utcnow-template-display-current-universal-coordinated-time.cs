using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph label
        builder.Writeln("Current UTC time:");

        // Insert a DATE field that will display the date and time
        // The field result will be formatted using the specified pattern
        FieldDate dateField = (FieldDate)builder.InsertField(FieldType.FieldDate, true);
        dateField.Format.DateTimeFormat = "yyyy-MM-dd HH:mm:ss";

        // Update the field so that it contains the current value
        doc.UpdateFields();

        // Prepare save options and set the custom time zone to UTC.
        // This makes the DATE field render its value using UTC rather than the local system time.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
        saveOptions.CustomTimeZoneInfo = TimeZoneInfo.Utc;

        // Save the document
        doc.Save("UtcDateTimeTemplate.docx", saveOptions);
    }
}
