using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add custom variables to the document.
        doc.Variables.Add("Company", "Acme Corp");
        doc.Variables.Add("Year", DateTime.Now.Year.ToString());

        // Insert a DOCVARIABLE field that will display the "Company" variable.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "Company";

        // Insert a line break and another DOCVARIABLE field for the "Year" variable.
        builder.Writeln();
        FieldDocVariable yearField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        yearField.VariableName = "Year";

        // Update all fields so they reflect the current variable values.
        doc.UpdateFields();

        // Configure MHTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for resources (images, fonts, CSS) in the MHTML file.
            ExportCidUrlsForMhtmlResources = true,
            // Export fonts as separate resources.
            ExportFontResources = true,
            // Store CSS in an external file.
            CssStyleSheetType = CssStyleSheetType.External,
            // Make the output HTML more readable.
            PrettyFormat = true
        };

        // Save the document as an MHTML file.
        doc.Save("DocumentWithVariables.mht", saveOptions);
    }
}
