using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Template.docx");

        // Define the merge fields and their corresponding values.
        string[] fieldNames = { "FullName", "Company", "Address", "City" };
        object[] fieldValues = { "James Bond", "MI5 Headquarters", "Milbank", "London" };

        // Perform a mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the resulting document as a PNG image (first page rendered).
        doc.Save("MergedOutput.png", SaveFormat.Png);
    }
}
