using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

class InsertOfficeMathFromLatex
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Writeln("Below is an equation inserted from a LaTeX string:");

        // Insert an empty EQ field. This field will later be converted to an OfficeMath object.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move the cursor to the field separator so we can write the LaTeX code.
        builder.MoveTo(eqField.Separator);
        // Example LaTeX string – you can replace this with any valid LaTeX expression.
        string latex = @"\frac{a}{b}";
        builder.Write(latex);

        // Return the cursor to the start of the field (optional, just for clarity).
        builder.MoveTo(eqField.Start);

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        // Insert the OfficeMath object before the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);

        // Remove the original EQ field, leaving only the OfficeMath node.
        eqField.Remove();

        // Optionally set the display type (Display = on its own line).
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        officeMath.Justification = OfficeMathJustification.Left;

        // Save the document.
        doc.Save("OfficeMathFromLatex.docx");
    }
}
