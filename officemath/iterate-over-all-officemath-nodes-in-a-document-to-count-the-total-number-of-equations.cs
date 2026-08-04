using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathCounter
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several EQ fields with simple equations.
        InsertFieldEQ(builder, @"\f(1,2)"); // Fraction 1/2
        InsertFieldEQ(builder, @"\r(3,x)"); // Cube root of x
        InsertFieldEQ(builder, @"\i \su(n=1,5,n)"); // Integral with summation
        InsertFieldEQ(builder, @"\s \up8(Sup) \s \do8(Sub)"); // Superscript and subscript

        // Convert each EQ field to a real OfficeMath node.
        var eqFields = doc.Range.Fields.OfType<FieldEQ>().ToList();
        foreach (FieldEQ eqField in eqFields)
        {
            OfficeMath officeMath = eqField.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
                // Remove the original field.
                eqField.Remove();
            }
        }

        // Count top‑level OfficeMath paragraph nodes (actual equations).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int equationCount = 0;
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                equationCount++;
        }

        // Output the result.
        Console.WriteLine($"Total number of equations: {equationCount}");

        // Save the document (optional, demonstrates that the file was created).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Equations.docx");
        doc.Save(outputPath, SaveFormat.Docx);
    }

    // Helper method to insert an EQ field with the specified arguments.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert the EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the equation arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);
        // Move back to the field start's parent node to continue building.
        builder.MoveTo(field.Start.ParentNode);
        // Insert a paragraph break after each equation for readability.
        builder.InsertParagraph();
        return field;
    }
}
