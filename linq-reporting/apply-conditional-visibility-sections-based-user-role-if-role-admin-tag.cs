using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section visible to every user.
        builder.Writeln("Public Section: Visible to all users.");

        // Insert an IF field that shows the following text only when the MERGEFIELD \"role\" equals \"Admin\".
        // The field code is:
        // IF MERGEFIELD role = "Admin" "Admin Section: Confidential content." ""
        builder.InsertField(@"IF MERGEFIELD role = ""Admin"" ""Admin Section: Confidential content."" """"");

        // Save the document.
        doc.Save("ConditionalVisibility.docx");
    }
}
