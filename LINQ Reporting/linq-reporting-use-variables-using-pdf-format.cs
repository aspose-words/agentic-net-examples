using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class LinqReportingWithVariablesPdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DOCVARIABLE field that will display the value of a document variable.
        // The field name is "MyVariable".
        FieldDocVariable docVariableField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        docVariableField.VariableName = "MyVariable";

        // Add a line break after the field for readability.
        builder.Writeln();

        // Define a document variable with the same name used in the field.
        doc.Variables.Add("MyVariable", "Hello Aspose.Words LINQ Reporting!");

        // The LINQ Reporting Engine can be used to process the template.
        // In this simple case we do not need an external data source, but we still invoke BuildReport
        // to ensure that any template syntax is evaluated.
        ReportingEngine engine = new ReportingEngine();
        // No data source is required; pass null and an empty name.
        engine.BuildReport(doc, null, "");

        // Prepare PDF save options (optional customizations can be set here).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: set the PDF to display the document title in the viewer.
            DisplayDocTitle = true
        };

        // Define the output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "LinqReportingWithVariables.pdf");

        // Save the document as PDF using the specified options.
        doc.Save(outputPath, pdfOptions);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
