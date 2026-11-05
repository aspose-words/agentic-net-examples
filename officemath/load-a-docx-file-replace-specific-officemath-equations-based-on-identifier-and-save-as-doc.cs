using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathReplaceExample
{
    public static void Main()
    {
        // Prepare temporary folder
        string folderPath = Path.Combine(Path.GetTempPath(), "OfficeMathExample");
        Directory.CreateDirectory(folderPath);

        // Paths for input and output documents
        string inputPath = Path.Combine(folderPath, "Input.docx");
        string outputPath = Path.Combine(folderPath, "Output.docx");

        // Create a sample document with two equations identified by bookmarks
        CreateSampleDocument(inputPath);

        // Load the document and replace equations based on bookmark identifiers
        Document doc = new Document(inputPath);
        ReplaceEquation(doc, "Eq1", @"\f(5,6)"); // Replace first equation with a new fraction
        ReplaceEquation(doc, "Eq2", @"\r(4,y)"); // Replace second equation with a new root

        // Save the modified document
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the output file exists and contains the expected number of equations
        ValidateOutput(outputPath);
    }

    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First equation with bookmark "Eq1"
        builder.StartBookmark("Eq1");
        InsertEquationViaEQField(builder, @"\f(1,2)"); // fraction 1/2
        builder.EndBookmark("Eq1");
        builder.Writeln(); // new line

        // Second equation with bookmark "Eq2"
        builder.StartBookmark("Eq2");
        InsertEquationViaEQField(builder, @"\r(3,x)"); // cube root of x
        builder.EndBookmark("Eq2");
        builder.Writeln();

        doc.Save(filePath, SaveFormat.Docx);
    }

    private static void InsertEquationViaEQField(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an EQ field
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        if (field is FieldEQ fieldEQ)
        {
            // Write the EQ argument string into the field separator
            DocumentBuilder sepBuilder = new DocumentBuilder(builder.Document);
            sepBuilder.MoveTo(fieldEQ.Separator);
            sepBuilder.Write(eqSwitch);

            // Update the field so that Aspose.Words can parse the EQ switch
            field.Update();

            // Convert the field to a real OfficeMath object
            OfficeMath officeMath = fieldEQ.AsOfficeMath();

            // If conversion succeeded, insert the OfficeMath node before the field start
            if (officeMath != null)
            {
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            }

            // Remove the original field (whether conversion succeeded or not)
            field.Remove();
        }
    }

    private static void ReplaceEquation(Document doc, string bookmarkName, string newEqSwitch)
    {
        Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];
        if (bookmark == null)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

        // Locate the paragraph that contains the bookmark start
        Paragraph para = bookmark.BookmarkStart.ParentNode as Paragraph;
        if (para == null)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' is not inside a paragraph.");

        // Find the first top‑level OfficeMath node inside the paragraph
        OfficeMath oldMath = para.GetChildNodes(NodeType.OfficeMath, false)
                                 .Cast<OfficeMath>()
                                 .FirstOrDefault();

        if (oldMath == null)
            throw new InvalidOperationException($"No OfficeMath found for bookmark '{bookmarkName}'.");

        // Move the builder to the location of the old equation
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(oldMath);

        // Remove the old OfficeMath node
        oldMath.Remove();

        // Insert the new equation using the deterministic EQ‑field workflow
        InsertEquationViaEQField(builder, newEqSwitch);
    }

    private static void ValidateOutput(string outputPath)
    {
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        Document resultDoc = new Document(outputPath);
        // Count only top‑level equations (OfficeMath nodes whose MathObjectType is OMathPara)
        int equationCount = resultDoc.GetChildNodes(NodeType.OfficeMath, true)
                                     .Cast<OfficeMath>()
                                     .Count(m => m.MathObjectType == MathObjectType.OMathPara);

        if (equationCount != 2)
            throw new InvalidOperationException($"Expected 2 top‑level equations after replacement, but found {equationCount}.");
    }
}
