using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class SetOfficeMathDisplayInline
{
    public static void Main()
    {
        // Output file path.
        const string outputPath = "OfficeMathInline.docx";

        // Create a new blank document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will be turned into a real OfficeMath object.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ argument (a simple fraction) after the field separator.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that its result is calculated (required for some versions).
        eqField.Update();

        // Return the builder to the start of the field (the field's parent paragraph).
        builder.MoveTo(eqField.Start);

        // Convert the EQ field to an OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Ensure we are working with a top‑level equation (MathObjectType.OMathPara).
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            // Set the display mode to Inline for a compact layout.
            officeMath.DisplayType = OfficeMathDisplayType.Inline;
        }

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
