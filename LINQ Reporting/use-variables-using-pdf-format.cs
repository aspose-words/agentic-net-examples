using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class PdfVariableExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add variables to the document's variable collection.
        VariableCollection vars = doc.Variables;
        vars.Add("Company", "Acme Corp.");
        vars.Add("ReportDate", DateTime.Now.ToString("D"));
        vars.Add("Author", "John Doe");

        // Use DocumentBuilder to insert DOCVARIABLE fields that will display the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the Company variable.
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "Company";
        companyField.Update();

        // Insert a line break and the ReportDate variable.
        builder.Writeln();
        FieldDocVariable dateField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        dateField.VariableName = "ReportDate";
        dateField.Update();

        // Insert another line break and the Author variable.
        builder.Writeln();
        FieldDocVariable authorField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        authorField.VariableName = "Author";
        authorField.Update();

        // Update all fields in the document (optional but ensures the values are refreshed).
        doc.UpdateFields();

        // Configure PDF save options (e.g., embed core fonts and set compliance level).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UseCoreFonts = true,                     // Substitute common TrueType fonts with PDF core fonts.
            Compliance = PdfCompliance.PdfA1b,       // Save as PDF/A-1b for long‑term preservation.
            PreserveFormFields = false               // No form fields needed for this example.
        };

        // Save the document as a PDF file using the configured options.
        doc.Save("DocumentWithVariables.pdf", pdfOptions);
    }
}
