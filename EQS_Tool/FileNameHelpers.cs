using System;

namespace EQS_Tool
{
    public class FileNameHelpers()
    {
        public static string MakeValidFileName(string fileName)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            foreach (var invalidChar in invalidChars)
            {
                fileName = fileName.Replace(invalidChar, '_');
            }
            return fileName;
        }

        public static string GenerateFileName(string certificateType, string[] certificateData, string namingFormat)
        {
            //function to generate file names based on certificate type and naming format
            //retuns [FileName, AlternateFileName]
            //required for moving files to correct folders and emailing

            // set variables from certificate data
            string uprn = certificateData[1];
            string date = certificateData[3];
            string result = certificateData[9];

            string namingConvention;
            string fileName;

            string[] dateArray = date.Split("/");
            uprn = uprn.ToUpper();

            if (certificateType == "EICR")
            {
                if (uprn.Contains("B") || uprn.Contains("C"))
                {
                    namingConvention = "CEICR";
                }
                else
                {
                    namingConvention = "DEICR";
                }

                if (result == "UNSAT")
                {
                    fileName = namingFormat
                        .Replace("UPRN", uprn)
                        .Replace("TYPE", namingConvention)
                        .Replace("DD", dateArray[0])
                        .Replace("MM", dateArray[1])
                        .Replace("YY", dateArray[2].Substring(2)) + "_UNSAT.pdf";
                }
                else
                {
                    fileName = namingFormat
                        .Replace("UPRN", uprn)
                        .Replace("TYPE", namingConvention)
                        .Replace("DD", dateArray[0])
                        .Replace("MM", dateArray[1])
                        .Replace("YY", dateArray[2].Substring(2)) + ".pdf";
                }
            }
            else
            {
                fileName = namingFormat
                        .Replace("UPRN", uprn)
                        .Replace("TYPE", certificateType)
                        .Replace("DD", dateArray[0])
                        .Replace("MM", dateArray[1])
                        .Replace("YY", dateArray[2].Substring(2)) + ".pdf";
            }

            return fileName;
        }
    }
}

