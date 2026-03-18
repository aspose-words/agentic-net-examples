using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DATE field. This field will later be updated using the custom culture.
        builder.InsertField(FieldType.FieldDate, true);

        // Tell the document to obtain the culture from the field code.
        doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;

        // Assign a custom culture provider that always returns the German culture.
        doc.FieldOptions.FieldUpdateCultureProvider = new GermanCultureProvider();

        // Set the field's locale to German (de-DE) so that the provider is invoked.
        FieldDate dateField = (FieldDate)doc.Range.Fields[0];
        dateField.LocaleId = new CultureInfo("de-DE").LCID;

        // Load CSV data that contains European date formats (e.g., "31.12.2023").
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ';'; // Example delimiter.
        CsvDataSource csvSource = new CsvDataSource("data.csv", loadOptions);

        // Example usage: perform a mail merge using a column named "Date".
        // The date value will be parsed according to the German culture supplied above.
        doc.MailMerge.Execute(new[] { "Date" }, new object[] { DateTime.Now });

        // Update fields so that the DATE field reflects the custom culture formatting.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("Output.docx");
    }

    // Custom culture provider that returns German culture for any field.
    private class GermanCultureProvider : IFieldUpdateCultureProvider
    {
        public CultureInfo GetCulture(string name, Field field)
        {
            // Ignore the requested culture name and always return German culture.
            return new CultureInfo("de-DE");
        }
    }
}
