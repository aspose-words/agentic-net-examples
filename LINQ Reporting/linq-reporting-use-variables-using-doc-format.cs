using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DOCVARIABLE field that will display the value of a document variable.
        FieldDocVariable varField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        varField.VariableName = "Greeting";

        // Add a variable to the document's variable collection.
        doc.Variables.Add("Greeting", "Hello, World!");

        // Use the LINQ Reporting Engine to process the template.
        // No external data source is required for this example.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, null, null);

        // Update all fields so the DOCVARIABLE field reflects the variable's value.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
