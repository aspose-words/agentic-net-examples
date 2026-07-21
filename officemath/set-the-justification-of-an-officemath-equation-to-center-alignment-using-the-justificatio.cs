using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will be converted to a real OfficeMath object.
        // The equation "\f(1,2)" creates a simple fraction 1/2.
        FieldEQ eqField = InsertFieldEQ(builder, @"\f(1,2)");

        // Update fields so that the EQ field is fully formed before conversion.
        doc.UpdateFields();

        // Ensure the field result is up‑to‑date.
        eqField.Update();

        // Convert the EQ field to OfficeMath.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Replace the field with the generated OfficeMath node.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Work only with top‑level OfficeMath (MathObjectType.OMathPara).
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            // Set display type to Display before changing justification (required by the API).
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            // Center the equation.
            officeMath.Justification = OfficeMathJustification.Center;
        }

        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "JustifiedEquation.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Reload the document and verify the justification.
        Document loadedDoc = new Document(outputPath);
        OfficeMath loadedMath = (OfficeMath)loadedDoc.GetChild(NodeType.OfficeMath, 0, true);
        if (loadedMath == null)
            throw new InvalidOperationException("No OfficeMath node found in the saved document.");

        if (loadedMath.Justification != OfficeMathJustification.Center)
            throw new InvalidOperationException("The OfficeMath justification was not set to Center.");
    }

    // Helper method to insert an EQ field with the specified arguments.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert the EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the separator and write the equation arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);
        // Return to the field's paragraph and start a new paragraph after it.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();
        return field;
    }
}
