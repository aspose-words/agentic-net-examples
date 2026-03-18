using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class InlineErrorDemo
{
    static void Main()
    {
        // Load a template that contains LINQ Reporting Engine syntax (e.g., <<[MissingField]>>).
        // The Document constructor both creates the object and loads the file.
        Document template = new Document("Template.docx");

        // Configure the ReportingEngine to inline syntax error messages.
        // The InlineErrorMessages flag makes the engine insert error details directly into the output.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report. An empty data source is supplied to provoke any syntax errors present in the template.
        engine.BuildReport(template, new object());

        // Save the resulting document. Errors will appear inline, e.g., <<error: ...>>.
        template.Save("Result.docx");
    }
}
