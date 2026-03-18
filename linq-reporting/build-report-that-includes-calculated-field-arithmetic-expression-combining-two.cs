using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Reporting;

class CalculatedReport
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a label for the calculated value.
        builder.Writeln("Total (value1 + value2):");

        // Build a MERGEFIELD for the first JSON property.
        FieldBuilder value1Field = new FieldBuilder(FieldType.FieldMergeField);
        value1Field.AddArgument("value1"); // <<[data.value1]>>

        // Build a MERGEFIELD for the second JSON property.
        FieldBuilder value2Field = new FieldBuilder(FieldType.FieldMergeField);
        value2Field.AddArgument("value2"); // <<[data.value2]>>

        // Build a formula field that adds the two MERGEFIELDs.
        FieldBuilder formula = new FieldBuilder(FieldType.FieldFormula);
        formula.AddArgument(value1Field);   // First operand.
        formula.AddArgument("+");           // Operator.
        formula.AddArgument(value2Field);   // Second operand.

        // Insert the formula field into the document.
        // It will be placed after the label we wrote above.
        formula.BuildAndInsert(doc.FirstSection.Body.LastParagraph);

        // Prepare JSON data source.
        string json = @"{ ""value1"": 10, ""value2"": 25 }";
        using (MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
        {
            JsonDataSource dataSource = new JsonDataSource(jsonStream);

            // Build the report, binding the JSON data source to the name "data".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "data");
        }

        // Update all fields so that the formula evaluates with the supplied data.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("CalculatedReport.docx");
    }
}
