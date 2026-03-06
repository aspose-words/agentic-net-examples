using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Words.Fields; // Added for FieldDocVariable and FieldType

class RtfVariableExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some document variables.
        doc.Variables.Add("Company", "Contoso Ltd.");
        doc.Variables.Add("Address", "123 Main St., Anytown");
        doc.Variables.Add("Phone", "+1 (555) 123‑4567");

        // Use a DocumentBuilder to insert DOCVARIABLE fields that display the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with the Company variable.
        builder.Writeln("Company: ");
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "Company";
        companyField.Update();

        // Insert a paragraph with the Address variable.
        builder.Writeln("Address: ");
        FieldDocVariable addressField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        addressField.VariableName = "Address";
        addressField.Update();

        // Insert a paragraph with the Phone variable.
        builder.Writeln("Phone: ");
        FieldDocVariable phoneField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        phoneField.VariableName = "Phone";
        phoneField.Update();

        // Configure RTF save options.
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            // Reduce file size (acceptable because we have no RTL text).
            ExportCompactSize = true,
            // Do not include extra keywords for very old readers to keep the file smaller.
            ExportImagesForOldReaders = false,
            // Ensure the format is RTF (default, but explicit for clarity).
            SaveFormat = SaveFormat.Rtf
        };

        // Save the document as RTF.
        string rtfPath = "DocumentWithVariables.rtf";
        doc.Save(rtfPath, saveOptions);

        // Load the saved RTF back using RtfLoadOptions (demonstrating load options usage).
        RtfLoadOptions loadOptions = new RtfLoadOptions
        {
            // Preserve the original encoding detection behavior.
            RecognizeUtf8Text = true
        };
        Document loadedDoc = new Document(rtfPath, loadOptions);

        // Update fields after loading to ensure they reflect current variable values.
        loadedDoc.UpdateFields();

        // Save the loaded document as a new RTF file to verify the round‑trip.
        loadedDoc.Save("DocumentWithVariables_RoundTrip.rtf", saveOptions);
    }
}
