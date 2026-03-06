using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class DocmVariableExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some document variables.
        doc.Variables.Add("CompanyName", "Contoso Ltd.");
        doc.Variables.Add("ReportDate", DateTime.Today.ToString("d"));

        // Use DocumentBuilder to insert DOCVARIABLE fields that display the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field for CompanyName.
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "CompanyName";
        companyField.Update();

        builder.Writeln(); // Add a line break.

        // Insert a field for ReportDate.
        FieldDocVariable dateField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        dateField.VariableName = "ReportDate";
        dateField.Update();

        // Save the document as a macro‑enabled DOCM file.
        doc.Save("VariablesExample.docm", SaveFormat.Docm);
    }
}
