using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class SetOfficeMathJustification
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will later be converted to a real OfficeMath object.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write a simple fraction switch at the field separator.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)"); // fraction 1/2

        // Return the cursor to the field start.
        builder.MoveTo(eqField.Start);

        // Update fields so the EQ field is evaluated and can be converted.
        doc.UpdateFields();

        // Retrieve the EQ field again (ensures the field is up‑to‑date).
        eqField = (FieldEQ)doc.Range.Fields.OfType<FieldEQ>().FirstOrDefault();
        if (eqField == null)
            throw new InvalidOperationException("EQ field not found after update.");

        // Convert the EQ field to an OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Replace the field with the real OfficeMath node.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Set display type before changing justification (required by API).
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        // Center the equation.
        officeMath.Justification = OfficeMathJustification.Center;

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OfficeMathCentered.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
