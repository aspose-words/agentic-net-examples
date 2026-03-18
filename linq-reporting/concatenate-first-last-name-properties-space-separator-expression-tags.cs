using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

class ConcatenateNameExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a formula field that concatenates the MERGEFIELD values for
        // "First Name" and "Last Name" with a space between them.
        // The field code looks like:
        // = { MERGEFIELD "First Name" } & " " & { MERGEFIELD "Last Name" }
        builder.InsertField("= { MERGEFIELD \"First Name\" } & \" \" & { MERGEFIELD \"Last Name\" }");

        // Prepare a data source for the mail merge.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("First Name");
        table.Columns.Add("Last Name");
        table.Rows.Add("John", "Doe");
        table.Rows.Add("Jane", "Smith");

        // Execute the mail merge – the formula field will evaluate to "John Doe", etc.
        doc.MailMerge.Execute(table);

        // Save the result.
        doc.Save("ConcatenatedNames.docx");
    }
}
