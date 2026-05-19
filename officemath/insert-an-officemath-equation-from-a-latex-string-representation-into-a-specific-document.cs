using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class InsertOfficeMathFromLatex
{
    public static void Main()
    {
        // LaTeX source (kept as comment for reference):
        // \[ x^{2} + y^{2} = z^{2} \]
        // Aspose.Words cannot parse LaTeX directly, so we create a simple
        // OfficeMath equation using the deterministic EQ‑field bootstrap workflow.

        // 1. Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Write introductory text.
        builder.Writeln("Example document with an inserted equation:");

        // 3. Insert an empty paragraph that will hold the equation.
        builder.Writeln();

        // 4. Insert an EQ field and write a safe EQ argument.
        // The argument "\f(1,2)" creates a simple fraction 1/2.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator to write the EQ switch arguments.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");
        // Update the field so that its internal representation becomes valid.
        eqField.Update();
        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);

        // 5. Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // 6. Insert the OfficeMath node before the field start and remove the original field.
        Node fieldStart = eqField.Start;
        ((CompositeNode)fieldStart.ParentNode).InsertBefore(officeMath, fieldStart);
        eqField.Remove();

        // 7. Set display formatting for the top‑level equation.
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        officeMath.Justification = OfficeMathJustification.Center;

        // 8. Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");
        doc.Save(outputPath);

        // 9. Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
