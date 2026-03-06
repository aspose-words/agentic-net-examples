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

        // Add variables using the collection's Add method.
        vars.Add("Company", "Acme Corp.");
        vars.Add("Year", "2026");

        // Alternatively, set variables directly via the string indexer (DOT format).
        // This will add the variable if it does not exist, or update it if it does.
        vars["Location"] = "New York";
        vars["Year"] = "2027"; // Update existing variable.

        // Insert a DOCVARIABLE field that displays the "Company" variable.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "Company";
        companyField.Update(); // Resolve the field result.

        // Insert another DOCVARIABLE field for the "Location" variable.
        FieldDocVariable locationField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        locationField.VariableName = "Location";
        locationField.Update();

        // Insert a third DOCVARIABLE field for the "Year" variable.
        FieldDocVariable yearField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        yearField.VariableName = "Year";
        yearField.Update();

        // Demonstrate retrieving a variable value via the indexer.
        string companyValue = vars["Company"]; // "Acme Corp."
        Console.WriteLine($"Company variable: {companyValue}");

        // Demonstrate retrieving a variable by index.
        string firstVariableValue = vars[0]; // Variables are sorted alphabetically; first is "Company".
        Console.WriteLine($"First variable (by index): {firstVariableValue}");

        // Save the document to disk.
        doc.Save("DocumentWithVariables.docx");
    }
}
