using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class LinqReportingWithVariables
{
    static void Main()
    {
        // 1. Create a blank document.
        Document doc = new Document();

        // 2. Use DocumentBuilder to insert a DOCVARIABLE field that will display a document variable.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the field and keep it un‑updated so we can set the variable later.
        FieldDocVariable varField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        varField.VariableName = "Greeting";

        // 3. Define the variable value in the document's variable collection.
        doc.Variables["Greeting"] = "Hello from Aspose.Words!";

        // 4. (Optional) Run the ReportingEngine – here we have no external data source,
        //    but invoking BuildReport demonstrates that the engine can work together with variables.
        ReportingEngine engine = new ReportingEngine();
        // No data source is required for this simple example; pass null.
        engine.BuildReport(doc, null);

        // 5. Save the document in WordML (XML) format.
        WordML2003SaveOptions saveOptions = new WordML2003SaveOptions
        {
            // Make the XML output human‑readable.
            PrettyFormat = true,
            // Ensure the correct save format is set (required by the options class).
            SaveFormat = SaveFormat.WordML
        };

        doc.Save("LinqReportingWithVariables.xml", saveOptions);
    }
}
