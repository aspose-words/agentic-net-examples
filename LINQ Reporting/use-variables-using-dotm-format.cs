using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document that will be saved as a DOTM (macro‑enabled template).
        Document doc = new Document();

        // Access the document's variable collection.
        VariableCollection variables = doc.Variables;

        // Add variables that can be referenced later by DOCVARIABLE fields.
        variables.Add("CompanyName", "Acme Corp");
        variables.Add("ReportDate", DateTime.Now.ToString("d"));

        // Use DocumentBuilder to insert fields that display the variable values.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DOCVARIABLE field for the CompanyName variable.
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "CompanyName";
        companyField.Update();

        // Add a line break between fields.
        builder.Writeln();

        // Insert a DOCVARIABLE field for the ReportDate variable.
        FieldDocVariable dateField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        dateField.VariableName = "ReportDate";
        dateField.Update();

        // Save the document as a DOTM file.
        doc.Save("Template.dotm", SaveFormat.Dotm);
    }
}
