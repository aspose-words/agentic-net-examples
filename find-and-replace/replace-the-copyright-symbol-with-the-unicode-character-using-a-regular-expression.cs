using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Example text with a copyright symbol: (c) 2023 Aspose.");

        // Define a regular expression that matches the copyright symbol written as "(c)".
        Regex copyrightRegex = new Regex(@"\(c\)", RegexOptions.IgnoreCase);

        // Perform the replacement with the Unicode © character.
        int replacedCount = doc.Range.Replace(copyrightRegex, "©", new FindReplaceOptions());

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one copyright symbol replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
