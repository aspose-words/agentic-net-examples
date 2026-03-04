using System;
using Aspose.Words;
using Aspose.Words.Fields;

class DocmVariableExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add variables to the document's variable collection.
        doc.Variables.Add("Company", "Acme Corp");
        doc.Variables.Add("Year", "2026");

        // Use DocumentBuilder to insert DOCVARIABLE fields that reference the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first variable field.
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "Company";
        companyField.Update();

        // Add a line break between fields.
        builder.Writeln();

        // Insert second variable field.
        FieldDocVariable yearField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        yearField.VariableName = "Year";
        yearField.Update();

        // Save the document as a macro‑enabled DOCM file.
        doc.Save("Variables.docm", SaveFormat.Docm);
    }
}
