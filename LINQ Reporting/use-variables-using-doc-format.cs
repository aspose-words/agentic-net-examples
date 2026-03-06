using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Add custom variables to the document's variable collection.
        doc.Variables.Add("Company", "Acme Corp");
        doc.Variables.Add("Address", "123 Main St, Springfield");

        // Use DocumentBuilder to insert DOCVARIABLE fields that display the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line for the Company variable.
        builder.Writeln("Company: ");
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "Company";
        companyField.Update(); // Ensure the field result reflects the variable value.

        // Insert a line for the Address variable.
        builder.Writeln("Address: ");
        FieldDocVariable addressField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        addressField.VariableName = "Address";
        addressField.Update();

        // Save the document in the legacy DOC format using DocSaveOptions.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        doc.Save("DocumentWithVariables.doc", saveOptions);
    }
}
