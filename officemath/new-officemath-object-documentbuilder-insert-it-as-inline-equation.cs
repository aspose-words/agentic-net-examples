using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field (the field that can be converted to OfficeMath).
        // The field is inserted but not yet populated with equation code.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move the builder to the field separator to write the equation arguments.
        // Here we create a simple inline equation "x+y".
        builder.MoveTo(eqField.Separator);
        builder.Write(@"x+y");

        // Return the cursor to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        // Insert the OfficeMath object inline before the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);

        // Remove the now‑redundant EQ field.
        eqField.Remove();

        // Save the document with the inline OfficeMath equation.
        doc.Save("InlineOfficeMath.docx");
    }
}
