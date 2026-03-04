using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Access the document's variable collection.
        VariableCollection vars = doc.Variables;

        // Add variables that will be referenced in the template.
        vars.Add("CompanyName", "Acme Corp.");
        vars.Add("ReportDate", DateTime.Today.ToString("yyyy-MM-dd"));
        vars.Add("Author", "John Doe");

        // Use DocumentBuilder to insert DOCVARIABLE fields for each variable.
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Company: ");
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "CompanyName";

        builder.Writeln("Date: ");
        FieldDocVariable dateField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        dateField.VariableName = "ReportDate";

        builder.Writeln("Prepared by: ");
        FieldDocVariable authorField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        authorField.VariableName = "Author";

        // Ensure the fields display the current variable values.
        companyField.Update();
        dateField.Update();
        authorField.Update();

        // Save the document as a DOTX template.
        doc.Save("Template.dotx", SaveFormat.Dotx);
    }
}
