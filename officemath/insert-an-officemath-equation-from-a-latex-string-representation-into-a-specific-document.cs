using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // LaTeX source string (metadata only, not directly parsed):
        // \frac{1}{2}
        const string latexEquation = @"\frac{1}{2}";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("Sample document with an inserted equation.");

        // Insert a paragraph that will hold the equation.
        builder.Writeln("The equation appears below:");

        // Insert an EQ field without updating it immediately.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, false);

        // Write a simple EQ switch that creates a fraction 1/2.
        // The switch \f(1,2) corresponds to the LaTeX \frac{1}{2}.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        // Ensure conversion succeeded before inserting.
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start node.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);

        // Remove the original EQ field from the document.
        eqField.Remove();

        // Set display formatting for the top‑level equation.
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        officeMath.Justification = OfficeMathJustification.Left;

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        // Validate that the document contains at least one OfficeMath node.
        int mathCount = doc.GetChildNodes(NodeType.OfficeMath, true).Count;
        if (mathCount == 0)
            throw new InvalidOperationException("No OfficeMath nodes were found in the saved document.");
    }
}
