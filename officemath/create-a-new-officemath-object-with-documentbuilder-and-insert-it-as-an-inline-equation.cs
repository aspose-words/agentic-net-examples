using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will later be converted to a real OfficeMath node.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write a simple EQ switch.
        // A leading space is required so the field code becomes "EQ \f(1,2)".
        builder.MoveTo(eqField.Separator);
        builder.Write(@" \f(1,2)");

        // Return the builder to the field's parent paragraph.
        builder.MoveTo(eqField.Start.ParentNode);

        // Ensure the field is up‑to‑date before conversion.
        eqField.Update();

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        // Verify conversion succeeded.
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Set the equation to be displayed inline.
        officeMath.DisplayType = OfficeMathDisplayType.Inline;

        // Save the document.
        doc.Save("OfficeMathInline.docx");
    }
}
