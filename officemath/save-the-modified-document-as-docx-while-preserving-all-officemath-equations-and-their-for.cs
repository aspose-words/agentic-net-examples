using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field (complex field) that will be converted to a real OfficeMath object.
        // The field is created with the FieldEquation type.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ argument.
        // The "\f(1,2)" switch creates a simple fraction 1/2.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);

        // Update the field so that its result is calculated before conversion.
        eqField.Update();

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start node.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        // Remove the original EQ field from the document.
        eqField.Remove();

        // Apply formatting to the top‑level OfficeMath node.
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        officeMath.Justification = OfficeMathJustification.Left;

        // Save the modified document as DOCX.
        string outputPath = "ModifiedDocument.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output DOCX file was not created.", outputPath);
    }
}
