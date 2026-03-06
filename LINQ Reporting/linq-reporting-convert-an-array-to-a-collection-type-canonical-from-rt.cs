using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source class containing an array.
    public class MyDataSource
    {
        // The array will be automatically treated as an IEnumerable by the ReportingEngine.
        public string[] Names { get; set; }

        public MyDataSource(string[] names)
        {
            Names = names;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the input RTF template and the output document.
            string inputRtfPath = Path.Combine(Environment.CurrentDirectory, "Template.rtf");
            string outputDocxPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // Load the RTF document with default load options.
            // RtfLoadOptions can be customized if needed (e.g., RecognizeUtf8Text).
            RtfLoadOptions loadOptions = new RtfLoadOptions();
            Document templateDoc = new Document(inputRtfPath, loadOptions);

            // Prepare the data source. The array will be converted to a collection
            // type internally by the ReportingEngine when the template accesses it.
            string[] nameArray = new[] { "Alice", "Bob", "Charlie" };
            MyDataSource data = new MyDataSource(nameArray);

            // Create the ReportingEngine and build the report.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "ds" can be used inside the template as <<[ds.Names]>>
            engine.BuildReport(templateDoc, data, "ds");

            // Save the populated document.
            templateDoc.Save(outputDocxPath, SaveFormat.Docx);
        }
    }
}
