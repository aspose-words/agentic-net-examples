using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains Aspose.Words reporting tags.
        Document template = new Document("Template.docx");

        // Prepare a data source object whose members will be referenced in the template.
        var data = new NumericTestClass
        {
            Value1 = 1234,
            Value2 = 5621718.589
        };

        // Create the reporting engine and populate the template with the data.
        ReportingEngine engine = new ReportingEngine();
        // The third argument ("ds") is the name used in the template to reference the data source.
        engine.BuildReport(template, data, "ds");

        // Save the populated document to a new DOCX file.
        template.Save("Result.docx");
    }
}

// Simple POCO class used as the data source for the report.
public class NumericTestClass
{
    public double Value1 { get; set; }
    public double Value2 { get; set; }
}
