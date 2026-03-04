using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the document's variable collection.
        VariableCollection vars = doc.Variables;

        // Add variables whose names contain dots.
        vars.Add("Customer.Name", "John Doe");
        vars.Add("Customer.Address", "123 Main St.");
        vars.Add("Order.Total", "99.99");

        // Use a DocumentBuilder to insert DOCVARIABLE fields that reference the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Customer Information:");
        InsertDocVariable(builder, "Customer.Name");
        builder.Writeln();
        InsertDocVariable(builder, "Customer.Address");
        builder.Writeln();

        builder.Writeln("Order Details:");
        InsertDocVariable(builder, "Order.Total");
        builder.Writeln();

        // Update all fields so they display the current variable values.
        doc.UpdateFields();

        // Save the document.
        doc.Save("VariablesDotFormat.docx");
    }

    // Helper method that inserts a DOCVARIABLE field for the specified variable name.
    static void InsertDocVariable(DocumentBuilder builder, string variableName)
    {
        // Insert a DOCVARIABLE field.
        FieldDocVariable field = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        // Assign the variable name (case‑insensitive).
        field.VariableName = variableName;
        // Update the field to reflect the variable's value.
        field.Update();
    }
}
