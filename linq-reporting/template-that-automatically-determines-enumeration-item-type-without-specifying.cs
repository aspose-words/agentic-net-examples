using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a reporting engine template that iterates over the ListTemplate enum.
        // The template does not need generic type parameters because the enum type
        // will be supplied to the engine via the KnownTypes collection.
        builder.Writeln("<<foreach [in ListTemplate]>>");
        builder.Writeln("<<[value]>>");
        builder.Writeln("<</foreach>>");

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register the ListTemplate enum as a known type.
        // This makes the enum accessible inside the template without any explicit data source.
        engine.KnownTypes.Add(typeof(ListTemplate));

        // Build the report. No external data source is required.
        engine.BuildReport(doc, null);

        // Save the generated document.
        doc.Save("EnumTemplateResult.docx");
    }
}
