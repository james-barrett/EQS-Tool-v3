using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;

namespace EQS_Tool
{
    public class CertificateFunctions
    {
        public static string[] CertificateData(string certificateType, Page[] pages, ApplicationConfig appConfig)
        {
            // get needed paths and fields from text bounding boxes
            // returns [JOB REFERENCE, UPRN, CERTIFICATE NUMBER, DATE, ADDRESS LINE 1, ADDRESS LINE 2, POSTCODE, ENGINEER, SUPERVISOR, RESULT, OCCUPIER]

            Page page_1 = pages[1];

            IEnumerable<Word> words;

            words = page_1.GetWords();
            string[] certificateData = new string[11];

            foreach (var word in words)
            {
                if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].Job.Left &&
                    word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].Job.Bottom &&
                    word.BoundingBox.Right <= appConfig.CertificateData[certificateType].Job.Right &&
                    word.BoundingBox.Top <= appConfig.CertificateData[certificateType].Job.Top)
                {
                    certificateData[0] += word;
                }
                else if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].UPRN.Left &&
                            word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].UPRN.Bottom &&
                            word.BoundingBox.Right <= appConfig.CertificateData[certificateType].UPRN.Right &&
                            word.BoundingBox.Top <= appConfig.CertificateData[certificateType].UPRN.Top)
                {
                    certificateData[1] += word;
                }
                else if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].Certificate.Left &&
                            word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].Certificate.Bottom &&
                            word.BoundingBox.Right <= appConfig.CertificateData[certificateType].Certificate.Right &&
                            word.BoundingBox.Top <= appConfig.CertificateData[certificateType].Certificate.Top)
                {
                    certificateData[2] += word;
                }
                else if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].Date.Left &&
                            word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].Date.Bottom &&
                            word.BoundingBox.Right <= appConfig.CertificateData[certificateType].Date.Right &&
                            word.BoundingBox.Top <= appConfig.CertificateData[certificateType].Date.Top)
                {
                    certificateData[3] += word;
                }
                else if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].Address1.Left &&
                            word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].Address1.Bottom &&
                            word.BoundingBox.Right <= appConfig.CertificateData[certificateType].Address1.Right &&
                            word.BoundingBox.Top <= appConfig.CertificateData[certificateType].Address1.Top)
                {
                    certificateData[4] += word + " ";
                }
                else if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].Address2.Left &&
                            word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].Address2.Bottom &&
                            word.BoundingBox.Right <= appConfig.CertificateData[certificateType].Address2.Right &&
                            word.BoundingBox.Top <= appConfig.CertificateData[certificateType].Address2.Top)
                {
                    certificateData[5] += word + " ";
                }
                else if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].PostCode.Left &&
                            word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].PostCode.Bottom &&
                            word.BoundingBox.Right <= appConfig.CertificateData[certificateType].PostCode.Right &&
                            word.BoundingBox.Top <= appConfig.CertificateData[certificateType].PostCode.Top)
                {
                    certificateData[6] += word + " ";
                }
                else if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].Engineer.Left &&
                            word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].Engineer.Bottom &&
                            word.BoundingBox.Right <= appConfig.CertificateData[certificateType].Engineer.Right &&
                            word.BoundingBox.Top <= appConfig.CertificateData[certificateType].Engineer.Top)
                {
                    certificateData[7] += word + " ";
                }
                else if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].Supervisor.Left &&
                            word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].Supervisor.Bottom &&
                            word.BoundingBox.Right <= appConfig.CertificateData[certificateType].Supervisor.Right &&
                            word.BoundingBox.Top <= appConfig.CertificateData[certificateType].Supervisor.Top)
                {
                    certificateData[8] += word + " ";
                }
                else if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].Occupier.Left &&
                            word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].Occupier.Bottom &&
                            word.BoundingBox.Right <= appConfig.CertificateData[certificateType].Occupier.Right &&
                            word.BoundingBox.Top <= appConfig.CertificateData[certificateType].Occupier.Top)
                {
                    certificateData[10] += word + " ";
                }
            }


            if (certificateType == "DFHN")
            {
                Page page_2 = pages[2];
                words = page_2.GetWords();

                foreach (var word in words)
                {
                    if (word.BoundingBox.Left >= appConfig.CertificateData[certificateType].Date.Left &&
                        word.BoundingBox.Bottom >= appConfig.CertificateData[certificateType].Date.Bottom &&
                        word.BoundingBox.Right <= appConfig.CertificateData[certificateType].Date.Right &&
                        word.BoundingBox.Top <= appConfig.CertificateData[certificateType].Date.Top)
                    {
                        certificateData[3] += word;
                    }
                }
            }



            return certificateData;
        }
        public static string CertificateType(string text)
        {
            var certificateTypes = new Dictionary<string, string>
                {
                    { "ELECTRICAL INSTALLATION CONDITION REPORT", "EICR" },
                    { "ELECTRICAL INSTALLATION CERTIFICATE", "EIC" },
                    { "MINOR ELECTRICAL INSTALLATION WORKS CERTIFICATE", "MW" },
                    { "DOMESTIC VISUAL CONDITION REPORT", "VIS" },
                    { "eNotification", "PARTP"},
                    { "COMMISSIONING OF A FIRE DETECTION", "DFHN" }
                };

            foreach (var certificate in certificateTypes)
            {
                if (text.Contains(certificate.Key))
                {
                    return certificate.Value;
                }
            }

            return "ERROR";
        }
        public static string CertificatePage(string text)
        {
            var certificatePages = new Dictionary<string, string>
                {
                    { "PART 1 : DETAILS", "CERTIFICATE_DETAILS" },
                    { "PART 5 : OBSERVATIONS", "OBSERVATIONS" },
                    { "PART 6 : DETAILS", "INSTALLATION_DETAILS" },
                    { "PART 9 : SCHEDULE", "SCHEDULE_INSPECTIONS" },
                    { "SCHEDULE OF CIRCUIT", "SCHEDULE_CIRCUIT" },
                    { "SCHEDULE OF TEST", "SCHEDULE_TEST" },
                    { "NOTES", "NOTES" },
                    { "NOTES FOR RECIPIENT", "NOTES FOR RECIPIENT" },
                    { "GUIDANCE FOR RECIPIENTS", "GUIDANCE FOR RECIPIENTS" }
                };

            foreach (var certificate in certificatePages)
            {
                if (text.Contains(certificate.Key))
                {
                    return certificate.Value;
                }
            }

            return "ERROR";
        }
        public static List<string> BuildFormattedLines(List<TextWithCoordinates> wordList)
        {
            var groupedText = wordList
                .OrderBy(t => t.Y1)
                .GroupBy(t => Math.Round(t.Y1 / 10))
                .Select(g => string.Join(" ", g.Select(t => t.Text)))
                .Select(line => line.Replace("(", "").Replace(")", ""))
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .ToList();
            return groupedText;
        }

        public static void CertificateDataCollection()
        {
            //placeholder for certificate data collection logic
        }

        //function to determin the result of the certificate type EICR Satisfactory or Unsatisfactory
        //only way to determine to to analyze the lines drawn on the PDF via the operations
        //this is a workaround as the PDF does not contain the result in text form
        //check if the line is within expected bounding boxes for the result
        //if not log error and return "ERROR"
        public static string EICRResult(Page page, ApplicationConfig appConfig)
        {
            var lines = new List<(PdfPoint Start, PdfPoint End)>();
            PdfPoint? currentPoint = null;

            foreach (var op in page.Operations)
            {
                var str = op.ToString();
                var parts = str.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length >= 3 &&
                    double.TryParse(parts[0], out double x) &&
                    double.TryParse(parts[1], out double y))
                {
                    var point = new PdfPoint(x, y);

                    if (parts[^1] == "m") // MoveTo
                    {
                        currentPoint = point;
                    }
                    else if (parts[^1] == "l") // LineTo
                    {
                        if (currentPoint != null)
                        {
                            lines.Add((currentPoint.Value, point));
                            currentPoint = point;
                        }
                    }
                    else if (parts[^1] == "S") // StrokePath
                    {
                        currentPoint = null;
                    }
                }
            }

            if (lines.Count == 0)
            {
                return "ERROR";
            }

            // Output lines
            foreach (var (start, end) in lines)
            {
                //Console.WriteLine($"Line from ({start.X}, {start.Y}) to ({end.X}, {end.Y})");

                if (start.X >= appConfig.CertificateData["EICR"].Result.Left && start.Y >= appConfig.CertificateData["EICR"].Result.Bottom &&
                    end.X <= appConfig.CertificateData["EICR"].Result.Right && end.Y <= appConfig.CertificateData["EICR"].Result.Top)
                {
                    if (appConfig.Settings.SatisfactoryCertificateLineStart == start.X &&
                         appConfig.Settings.SatisfactoryCertificateLineEnd == end.X)
                    {
                        return "SAT";
                    }
                    else if (appConfig.Settings.UnsatisfactoryCertificateLineStart == start.X &&
                             appConfig.Settings.UnsatisfactoryCertificateLineEnd == end.X)
                    {
                        return "UNSAT";
                    }
                    else
                    {
                        return "ERROR";
                    }
                }
                else
                {
                    return "ERROR";
                }
            }

            return "ERROR";
        }

        public static void GetObservations(Page page)
        {
            var words = page.GetWords();

            if (words == null || !words.Any())
            {
                return;
            }

            var wordList = words.Select(word => new TextWithCoordinates
            {
                Text = word.Text,
                X1 = word.BoundingBox.BottomLeft.X,
                Y1 = word.BoundingBox.BottomLeft.Y,
                X2 = word.BoundingBox.TopRight.X,
                Y2 = word.BoundingBox.TopRight.Y
            }).ToList();

            var formattedLines = BuildFormattedLines(wordList);

            foreach (var line in formattedLines)
            {
                if (line.Contains("C3"))
                {
                    Console.WriteLine($"C3 ITEMS: {line}");
                }

                if (line.Contains("C2"))
                {
                    Console.WriteLine($"C2 ITEMS: {line}");
                }

                if (line.Contains("C1"))
                {
                    Console.WriteLine($"C1 ITEMS: {line}");
                }

                if (line.Contains("N/A"))
                {
                    Console.WriteLine($"N/A ITEMS: {line}");
                }

                Console.WriteLine(line);
            }

        }

    }
}
