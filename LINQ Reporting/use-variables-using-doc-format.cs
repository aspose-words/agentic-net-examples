using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some document variables.
        doc.Variables.Add("Company", "Contoso Ltd.");
        doc.Variables.Add("Address", "123 Main St., Anytown");

        // Use DocumentBuilder to insert DOCVARIABLE fields that display the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field for the "Company" variable.
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "Company";
        companyField.Update();

        // Insert a paragraph break.
        builder.Writeln();

        // Insert a field for the "Address" variable.
        FieldDocVariable addressField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        addressField.VariableName = "Address";
        addressField.Update();

        // Save the document in the legacy DOC format using DocSaveOptions.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        doc.Save("DocumentWithVariables.doc", saveOptions);
    }
}
