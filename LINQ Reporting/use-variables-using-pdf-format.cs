using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Access the document's variable collection.
        VariableCollection vars = doc.Variables;

        // Add variables that will be used in the document.
        vars.Add("Company", "Acme Corp.");
        vars.Add("Address", "123 Business Rd.");
        vars.Add("Year", "2024");

        // Use DocumentBuilder to insert DOCVARIABLE fields that display the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the Company variable.
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "Company";
        companyField.Update();

        builder.Writeln(); // New line.

        // Insert the Address variable.
        FieldDocVariable addressField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        addressField.VariableName = "Address";
        addressField.Update();

        builder.Writeln(); // New line.

        // Insert the Year variable.
        FieldDocVariable yearField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        yearField.VariableName = "Year";
        yearField.Update();

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Replace common TrueType fonts with core PDF Type 1 fonts.
            UseCoreFonts = true,
            // Save the PDF as PDF/A-1b for archival compliance.
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as a PDF file using the specified options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
