using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Data;

class Program
{
    static void Main()
    {
        // Load the DOCM template that contains the heading for the LINQ Reporting Engine introduction.
        Document template = new Document("ReportingEngineTemplate.docm");

        // Create an instance of the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. In this simple scenario we do not need any external data,
        // so we pass an empty DataSet as the data source.
        DataSet emptyDataSource = new DataSet();
        engine.BuildReport(template, emptyDataSource);

        // Save the populated document to the desired output format.
        template.Save("ReportingEngineIntroduction.docx");
    }
}
