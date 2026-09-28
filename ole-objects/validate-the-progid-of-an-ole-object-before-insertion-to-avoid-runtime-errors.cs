using System;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        string[] progIds = { "Excel.Application", "Word.Application", "Invalid.ProgID" };

        foreach (var progId in progIds)
        {
            Type comType = Type.GetTypeFromProgID(progId);
            if (comType == null)
            {
                Console.WriteLine($"ProgID '{progId}' is not registered.");
                continue;
            }

            try
            {
                object comObject = Activator.CreateInstance(comType);
                Console.WriteLine($"ProgID '{progId}' is valid and instance created.");
                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ProgID '{progId}' is registered but failed to instantiate: {ex.Message}");
            }
        }
    }
}
