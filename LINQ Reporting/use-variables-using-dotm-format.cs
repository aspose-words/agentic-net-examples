using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document that will serve as a DOTM template.
        Document doc = new Document();

        // Add variables to the document's variable collection.
        VariableCollection vars = doc.Variables;
        vars.Add("CompanyName", "Contoso Ltd.");
        vars.Add("ReportDate", DateTime.Today.ToString("d"));
        vars.Add("Author", "John Doe");

        // Insert DOCVARIABLE fields that display the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);
        InsertVariableField(builder, "CompanyName");
        builder.Writeln();
        InsertVariableField(builder, "ReportDate");
        builder.Writeln();
        InsertVariableField(builder, "Author");

        // Ensure all fields reflect the current variable values.
        doc.UpdateFields();

        // Save the document as a macro‑enabled template (.dotm).
        doc.Save("Template.dotm", SaveFormat.Dotm);
    }

    // Helper method to insert a DOCVARIABLE field for a given variable name.
    static void InsertVariableField(DocumentBuilder builder, string variableName)
    {
        FieldDocVariable field = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        field.VariableName = variableName;
        field.Update();
    }
}
