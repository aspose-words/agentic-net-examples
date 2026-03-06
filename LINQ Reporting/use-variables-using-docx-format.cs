using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Obtain the document's variable collection.
        VariableCollection variables = doc.Variables;

        // Add some variables that we want to display in the document.
        variables.Add("CompanyName", "Acme Corp.");
        variables.Add("ReportDate", DateTime.Today.ToString("MMMM d, yyyy"));
        variables.Add("TotalSales", "12345.67");

        // Use DocumentBuilder to insert DOCVARIABLE fields that reference the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with the company name.
        builder.Writeln("Company: ");
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "CompanyName";
        companyField.Update();

        // Insert a paragraph with the report date.
        builder.Writeln();
        builder.Writeln("Date: ");
        FieldDocVariable dateField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        dateField.VariableName = "ReportDate";
        dateField.Update();

        // Insert a paragraph with the total sales.
        builder.Writeln();
        builder.Writeln("Total Sales: $");
        FieldDocVariable salesField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        salesField.VariableName = "TotalSales";
        salesField.Update();

        // Save the document in DOCX format.
        doc.Save("VariablesDemo.docx");
    }
}
