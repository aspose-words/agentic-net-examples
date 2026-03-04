using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Add custom variables to the document's variable collection.
        doc.Variables.Add("Company", "Acme Corp");
        doc.Variables.Add("Address", "123 Main St, Metropolis");
        doc.Variables.Add("Phone", "+1 (555) 123‑4567");

        // Use DocumentBuilder to insert fields that will display the variable values.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DOCVARIABLE field for each variable and update its result.
        InsertVariableField(builder, "Company");
        InsertVariableField(builder, "Address");
        InsertVariableField(builder, "Phone");

        // Save the document in DOCX format.
        doc.Save("VariablesDemo.docx");
    }

    // Inserts a DOCVARIABLE field for the specified variable name, updates it, and adds a line break.
    static void InsertVariableField(DocumentBuilder builder, string variableName)
    {
        // Insert a DOCVARIABLE field; the 'true' argument updates the field immediately.
        FieldDocVariable field = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        field.VariableName = variableName;
        field.Update(); // Ensure the field result reflects the current variable value.

        // Add a paragraph break after the field for readability.
        builder.Writeln();
    }
}
