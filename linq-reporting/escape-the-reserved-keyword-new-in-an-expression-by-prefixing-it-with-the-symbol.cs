using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Create a blank Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting tag that references the property named "new".
            // The property name is "new" (escaped with @ only in C# code), so the tag uses model.new.
            builder.Writeln("Escaped keyword value: <<[model.new]>>");

            // Prepare the data model. The property is named "new" and is escaped with @ in C#.
            Model model = new Model
            {
                @new = "Hello from escaped keyword"
            };

            // Build the report using the model as the root data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            doc.Save("EscapedKeywordReport.docx");
        }
    }

    // Data model with a property whose identifier is the reserved keyword "new".
    public class Model
    {
        // The @ prefix allows the use of the reserved word as a property name.
        public string @new { get; set; } = string.Empty;
    }
}
