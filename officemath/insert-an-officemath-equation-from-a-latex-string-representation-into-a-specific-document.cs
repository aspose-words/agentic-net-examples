using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class InsertOfficeMathFromLatex
{
    public static void Main()
    {
        // Output file name.
        const string outputPath = "Output.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory text.
        builder.Writeln("Document before equation.");

        // Bookmark that marks the insertion point.
        builder.StartBookmark("InsertHere");
        builder.Writeln("Placeholder for equation.");
        builder.EndBookmark("InsertHere");

        // Move the builder to the bookmark.
        builder.MoveToBookmark("InsertHere");

        // Start a new paragraph for the equation.
        builder.Writeln();

        // Insert an EQ field – the required bootstrap for a real OfficeMath node.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ argument. The LaTeX source is kept as a comment.
        // LaTeX source (for reference): \frac{1}{2}
        // Using a safe EQ switch that reliably converts: \f(1,2)
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that its result is calculated before conversion.
        eqField.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Set display formatting for the top‑level equation.
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        officeMath.Justification = OfficeMathJustification.Center;

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        // Verify that the document now contains at least one OfficeMath node.
        int mathCount = doc.GetChildNodes(NodeType.OfficeMath, true).Count;
        if (mathCount == 0)
            throw new InvalidOperationException("No OfficeMath node was found in the saved document.");

        Console.WriteLine($"Document saved to '{outputPath}' with {mathCount} OfficeMath node(s).");
    }
}
