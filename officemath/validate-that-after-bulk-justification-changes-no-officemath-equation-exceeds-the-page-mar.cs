using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

public class OfficeMathJustificationValidator
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");                     // Simple fraction 1/2
        InsertEquation(builder, @"\r(3,x)");                     // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)");            // Integral with summation
        InsertEquation(builder, @"\s \up8(Sup) \s \do8(Sub)"); // Superscript and subscript

        // Ensure equations are displayed on their own lines.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                om.DisplayType = OfficeMathDisplayType.Display;
        }

        // Apply a bulk justification change to all top‑level OfficeMath paragraphs.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                om.Justification = OfficeMathJustification.CenterGroup;
        }

        // Validate that no equation exceeds the usable page width (page width minus left/right margins).
        double usableWidth = doc.FirstSection.PageSetup.PageWidth -
                             doc.FirstSection.PageSetup.LeftMargin -
                             doc.FirstSection.PageSetup.RightMargin;

        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType != MathObjectType.OMathPara)
                continue; // Skip nested math objects.

            // Render the equation to obtain its size.
            OfficeMathRenderer renderer = om.GetMathRenderer();
            double equationWidth = renderer.SizeInPoints.Width;

            if (equationWidth > usableWidth)
                throw new InvalidOperationException(
                    $"Equation exceeds page margins. Width: {equationWidth} pts, Usable: {usableWidth} pts.");
        }

        // Save the resulting document.
        const string outputPath = "OfficeMathJustificationValidated.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Indicate successful validation.
        Console.WriteLine("All OfficeMath equations fit within the page margins.");
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node, and removes the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field so that the EQ code is processed.
        field.Update();

        // Return the builder to the field's start parent node to continue building.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }
        else
        {
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");
        }

        // Add a paragraph break after the equation for readability.
        builder.Writeln();
    }
}
