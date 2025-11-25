using System;

namespace EQS_Tool
{
    public class Helpers
    {
        public static void CleanExit()
        {
            Console.WriteLine("Press any key to exit.....");
            Console.ReadKey();
            Environment.Exit(-1);
        }
    }
}