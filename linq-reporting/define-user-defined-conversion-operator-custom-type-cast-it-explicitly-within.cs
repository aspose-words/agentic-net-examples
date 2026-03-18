using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace MyNamespace
{
    // Custom type with explicit conversion operators.
    public class MyNumber
    {
        public double Value { get; }

        public MyNumber(double value) => Value = value;

        // Explicit conversion from string to MyNumber.
        public static explicit operator MyNumber(string s) => new MyNumber(double.Parse(s));

        // Explicit conversion from MyNumber to int.
        public static explicit operator int(MyNumber n) => (int)n.Value;
    }
}

// Data source class.
public class Data
{
    public string NumberString { get; set; }
}

class Program
{
    static void Main()
    {
        // Create a new document and insert a template placeholder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The template casts the string to MyNumber, then to int.
        builder.Writeln("Result: <<(int)(MyNamespace.MyNumber)[NumberString]>>");

        // Prepare the data object.
        var data = new Data { NumberString = "42.7" };

        // Set up the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register the custom type so it can be used in the template.
        engine.KnownTypes.Add(typeof(MyNamespace.MyNumber));

        // Build the report using the data source.
        engine.BuildReport(doc, data);

        // Save the generated document.
        doc.Save("Result.docx");
    }
}
