using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add custom document variables.
        doc.Variables.Add("Company", "Aspose");
        doc.Variables.Add("Address", "123 Main St.");

        // Insert a DOCVARIABLE field that displays the "Company" variable.
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "Company";
        companyField.Update();

        builder.Writeln(); // Move to a new paragraph.

        // Insert a DOCVARIABLE field that displays the "Address" variable.
        FieldDocVariable addressField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        addressField.VariableName = "Address";
        addressField.Update();

        // Save the document in WordML (XML) format with pretty formatting.
        WordML2003SaveOptions saveOptions = new WordML2003SaveOptions
        {
            PrettyFormat = true // Makes the XML output easier to read.
        };
        doc.Save("Variables.doc", saveOptions);
    }
}
