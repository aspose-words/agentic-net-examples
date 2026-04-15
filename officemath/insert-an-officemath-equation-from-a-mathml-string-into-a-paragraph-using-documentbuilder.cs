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

        // Add introductory paragraph.
        builder.Writeln("Below is an equation inserted from a MathML string:");

        // Original MathML (kept as comment):
        // <math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>

        // Insert an EQ field using a deterministic switch that reliably converts to OfficeMath.
        // Here we use a simple fraction 1/2 as a placeholder.
        FieldEQ eqField = InsertFieldEQ(builder, @"\f(1,2)");

        // Ensure the field is up‑to‑date before conversion.
        eqField.Update();

        // Convert the EQ field to a real OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Replace the field with the generated OfficeMath node.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Apply display formatting to the top‑level equation.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Center;
        }

        // Save the document.
        string outputPath = "InsertedEquation.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    // Helper that inserts an EQ field and writes its arguments.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string arguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(arguments);

        // Return the cursor to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);
        // Add a paragraph break after the field for readability.
        builder.InsertParagraph();

        return field;
    }
}
