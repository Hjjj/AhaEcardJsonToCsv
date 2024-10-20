using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using CsvHelper;
using Newtonsoft.Json.Linq;
using System.Reflection;
using System.Configuration;
using System.Drawing;

namespace AhaEcardJsonToCsv
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"Begin at {DateTime.Now}");

            string jsonDirectoryPath = ConfigurationManager.AppSettings["jsonFilesDirectory"];
            string csvFileName = ConfigurationManager.AppSettings["csvFileDirectory"] + "\\" + GetCsvFileName();
            List<eCardCsvSheet> eCardCSVList = new List<eCardCsvSheet>();

            foreach (string jsonFilePath in Directory.EnumerateFiles(jsonDirectoryPath, "*.json"))
            {
                try 
                {
                    JObject jsonObject = ReadJsonFile(jsonFilePath);
                    var linesOnTheCard = ConvertFromJsonToLineList(jsonObject);

                    var cardType = DetectImageCertType(linesOnTheCard);
                    if (cardType == eCardImageType.UNKNOWN)
                    {
                        continue;
                    }

                    eCardCsvSheet csvObject = BuildEcardCsvSheet(jsonObject, linesOnTheCard, cardType, jsonFilePath);
                    eCardCSVList.Add(csvObject);

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading {jsonFilePath}, {ex.Message}");
                    continue;
                }
            }

            try
            {
                using (var writer = new StreamWriter(csvFileName))
                {

                    var config = new CsvHelper.Configuration.CsvConfiguration(System.Globalization.CultureInfo.InvariantCulture)
                    {
                        Delimiter = "\t" // Set the delimiter to a tab character
                    };

                    using (var csv = new CsvWriter(writer, config))
                    {
                        csv.WriteRecords(eCardCSVList);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing to {csvFileName}, {ex.Message}");
            }

            Console.WriteLine($"Complete at {DateTime.Now}");
            Console.ReadKey();
        }

        private static eCardCsvSheet BuildEcardCsvSheet(
            JObject jsonObject, 
            List<string> eCardCSVList, 
            eCardImageType cardType, 
            string imageFileName)
        {
            try
            {

                switch (cardType)
                {
                    case eCardImageType.ACLS_PROVIDER_DOUBLE_SQUARE_68:
                        return new eCardCsvSheet
                        {
                            Filename = GetFileNameWithoutExtension(imageFileName),
                            LineCount = eCardCSVList.Count.ToString(),
                            WHRatio = getWHRatio(jsonObject).ToString(),
                            DupePct = CalculateDuplicatePercentage(jsonObject).ToString(),
                            URL = FindImageFile(imageFileName, ConfigurationManager.AppSettings["ecardImageDirectory"]),
                            CertTitle = eCardCSVList[0],
                            LeftTitle = $"{eCardCSVList[2]} {eCardCSVList[6]}",
                            FullName = eCardCSVList[12],
                            IssueDate = eCardCSVList[24],
                            RenewByDate = eCardCSVList[25],
                            TrainingCenterName = eCardCSVList[32],
                            InstructorName = eCardCSVList[33],
                            TrainingCenterId = eCardCSVList[41],
                            InstructorId = eCardCSVList[38],
                            TrainingCenterCityState = eCardCSVList[49],
                            EcardCode = eCardCSVList[46],
                            TrainingCenterPhoneNumber = eCardCSVList[58],
                            TrainingSiteName = eCardCSVList[62]
                        };
                    case eCardImageType.ACLS_PROVIDER_SINGLE_SQUARE_33:
                        return new eCardCsvSheet
                        {
                            Filename = GetFileNameWithoutExtension(imageFileName),
                            LineCount = eCardCSVList.Count.ToString(),
                            WHRatio = getWHRatio(jsonObject).ToString(),
                            DupePct = CalculateDuplicatePercentage(jsonObject).ToString(),
                            URL = FindImageFile(imageFileName, ConfigurationManager.AppSettings["ecardImageDirectory"]),
                            CertTitle = eCardCSVList[0],
                            LeftTitle = $"{eCardCSVList[1]} {eCardCSVList[2]}",
                            FullName = eCardCSVList[17],
                            IssueDate = eCardCSVList[12],
                            RenewByDate = eCardCSVList[13],
                            TrainingCenterName = eCardCSVList[16],
                            InstructorName = eCardCSVList[17],
                            TrainingCenterId = eCardCSVList[20],
                            InstructorId = eCardCSVList[21],
                            TrainingCenterCityState = eCardCSVList[24],
                            EcardCode = eCardCSVList[25],
                            TrainingCenterPhoneNumber = eCardCSVList[29],
                        };
                    case eCardImageType.ACLS_INSTRUCTOR_SINGLE_SQUARE_29:
                        return new eCardCsvSheet
                        {
                            Filename = GetFileNameWithoutExtension(imageFileName),
                            LineCount = eCardCSVList.Count.ToString(),
                            WHRatio = getWHRatio(jsonObject).ToString(),
                            DupePct = CalculateDuplicatePercentage(jsonObject).ToString(),
                            URL = FindImageFile(imageFileName, ConfigurationManager.AppSettings["ecardImageDirectory"]),
                            CertTitle = eCardCSVList[0],
                            LeftTitle = $"{eCardCSVList[1]} {eCardCSVList[3]}",
                            FullName = eCardCSVList[6],
                            IssueDate = eCardCSVList[13],
                            RenewByDate = eCardCSVList[14],
                            //TrainingCenterName = eCardCSVList[16],
                            //InstructorName = eCardCSVList[17],
                            TrainingCenterId = eCardCSVList[22],
                            InstructorId = eCardCSVList[18],
                            TrainingCenterCityState = eCardCSVList[24],
                            EcardCode = eCardCSVList[21],
                            TrainingCenterPhoneNumber = eCardCSVList[26],
                        };
                    case eCardImageType.ACLS_INSTRUCTOR_DOUBLE_SQUARE_INSTRUCTIONS_68:
                        return new eCardCsvSheet
                        {
                            Filename = GetFileNameWithoutExtension(imageFileName),
                            LineCount = eCardCSVList.Count.ToString(),
                            WHRatio = getWHRatio(jsonObject).ToString(),
                            DupePct = CalculateDuplicatePercentage(jsonObject).ToString(),
                            URL = FindImageFile(imageFileName, ConfigurationManager.AppSettings["ecardImageDirectory"]),
                            CertTitle = eCardCSVList[0],
                            LeftTitle = $"{eCardCSVList[4]} {eCardCSVList[8]}",
                            FullName = eCardCSVList[14],
                            IssueDate = eCardCSVList[46],
                            RenewByDate = eCardCSVList[47],
                            //TrainingCenterName = eCardCSVList[16],
                            //InstructorName = eCardCSVList[17],
                            TrainingCenterId = eCardCSVList[17],
                            InstructorId = eCardCSVList[42],
                            TrainingCenterCityState = eCardCSVList[27],
                            EcardCode = eCardCSVList[48],
                            TrainingCenterPhoneNumber = eCardCSVList[35],
                        };
                    case eCardImageType.ACLS_INSTRUCTOR_SINGLE_WALLET_INSTRUCTIONS_36:
                        return new eCardCsvSheet
                        {
                            Filename = GetFileNameWithoutExtension(imageFileName),
                            LineCount = eCardCSVList.Count.ToString(),
                            WHRatio = getWHRatio(jsonObject).ToString(),
                            DupePct = CalculateDuplicatePercentage(jsonObject).ToString(),
                            URL = FindImageFile(imageFileName, ConfigurationManager.AppSettings["ecardImageDirectory"]),
                            CertTitle = eCardCSVList[0],
                            LeftTitle = $"{eCardCSVList[2]} {eCardCSVList[4]}",
                            FullName = eCardCSVList[10],
                            IssueDate = eCardCSVList[24],
                            RenewByDate = eCardCSVList[25],
                            //TrainingCenterName = eCardCSVList[16],
                            //InstructorName = eCardCSVList[17],
                            TrainingCenterId = eCardCSVList[13],
                            InstructorId = eCardCSVList[28],
                            //TrainingCenterCityState = eCardCSVList[27],
                            EcardCode = eCardCSVList[26],
                            TrainingCenterPhoneNumber = eCardCSVList[20],
                        };
                    /////
                    case eCardImageType.BLS_SINGLE_SQUARE_34:
                        return new eCardCsvSheet
                        {
                            Filename = GetFileNameWithoutExtension(imageFileName),
                            LineCount = eCardCSVList.Count.ToString(),
                            WHRatio = getWHRatio(jsonObject).ToString(),
                            DupePct = CalculateDuplicatePercentage(jsonObject).ToString(),
                            URL = FindImageFile(imageFileName, ConfigurationManager.AppSettings["ecardImageDirectory"]),
                            CertTitle = eCardCSVList[0],
                            LeftTitle = $"{eCardCSVList[1]} {eCardCSVList[3]}",
                            FullName = eCardCSVList[6],
                            IssueDate = eCardCSVList[12],
                            RenewByDate = eCardCSVList[13],
                            TrainingCenterName = eCardCSVList[16],
                            InstructorName = eCardCSVList[17],
                            TrainingCenterId = eCardCSVList[20],
                            InstructorId = eCardCSVList[21],
                            TrainingCenterCityState = eCardCSVList[24],
                            EcardCode = eCardCSVList[25],
                            TrainingCenterPhoneNumber = eCardCSVList[29],
                        };
                    case eCardImageType.BLS_DOUBLE_SQUARE_66:
                        return new eCardCsvSheet
                        {
                            Filename = GetFileNameWithoutExtension(imageFileName),
                            LineCount = eCardCSVList.Count.ToString(),
                            WHRatio = getWHRatio(jsonObject).ToString(),
                            DupePct = CalculateDuplicatePercentage(jsonObject).ToString(),
                            URL = FindImageFile(imageFileName, ConfigurationManager.AppSettings["ecardImageDirectory"]),
                            CertTitle = eCardCSVList[0],
                            LeftTitle = $"{eCardCSVList[2]} {eCardCSVList[6]}",
                            FullName = eCardCSVList[12],
                            IssueDate = eCardCSVList[24],
                            RenewByDate = eCardCSVList[25],
                            TrainingCenterName = eCardCSVList[32],
                            InstructorName = eCardCSVList[33],
                            TrainingCenterId = eCardCSVList[40],
                            InstructorId = eCardCSVList[41],
                            TrainingCenterCityState = eCardCSVList[47],
                            EcardCode = eCardCSVList[48],
                            TrainingCenterPhoneNumber = eCardCSVList[58],
                        };
                    case eCardImageType.PALS_SINGLE_SQUARE_38:
                        return new eCardCsvSheet
                        {
                            Filename = GetFileNameWithoutExtension(imageFileName),
                            LineCount = eCardCSVList.Count.ToString(),
                            WHRatio = getWHRatio(jsonObject).ToString(),
                            DupePct = CalculateDuplicatePercentage(jsonObject).ToString(),
                            URL = FindImageFile(imageFileName, ConfigurationManager.AppSettings["ecardImageDirectory"]),
                            CertTitle = eCardCSVList[0],
                            LeftTitle = $"{eCardCSVList[1]} {eCardCSVList[5]}",
                            FullName = eCardCSVList[11],
                            IssueDate = eCardCSVList[17],
                            RenewByDate = eCardCSVList[18],
                            TrainingCenterName = eCardCSVList[21],
                            InstructorName = eCardCSVList[23],
                            TrainingCenterId = eCardCSVList[26],
                            InstructorId = eCardCSVList[27],
                            TrainingCenterCityState = eCardCSVList[30],
                            EcardCode = eCardCSVList[31],
                            TrainingCenterPhoneNumber = eCardCSVList[35],
                        };
                    case eCardImageType.PALS_DOUBLE_SQUARE_70:
                        return new eCardCsvSheet
                        {
                            Filename = GetFileNameWithoutExtension(imageFileName),
                            LineCount = eCardCSVList.Count.ToString(),
                            WHRatio = getWHRatio(jsonObject).ToString(),
                            DupePct = CalculateDuplicatePercentage(jsonObject).ToString(),
                            URL = FindImageFile(imageFileName, ConfigurationManager.AppSettings["ecardImageDirectory"]),
                            CertTitle = eCardCSVList[0],
                            LeftTitle = $"{eCardCSVList[2]} {eCardCSVList[8]}",
                            FullName = eCardCSVList[18],
                            IssueDate = eCardCSVList[30],
                            RenewByDate = eCardCSVList[31],
                            TrainingCenterName = eCardCSVList[38],
                            InstructorName = eCardCSVList[39],
                            TrainingCenterId = eCardCSVList[48],
                            InstructorId = eCardCSVList[47],
                            TrainingCenterCityState = eCardCSVList[56],
                            EcardCode = eCardCSVList[54],
                            TrainingCenterPhoneNumber = eCardCSVList[64],
                        };
                    default:
                        Console.WriteLine("default case");
                        return null;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error Building eCard Object from Lines {imageFileName}, {ex.Message}");
                return null;
            }
        }

        private static string FindImageFile(string jsonFilePath, string imageDirectoryPath)
        {
            // Step 1: Extract the filename without extension
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(jsonFilePath);

            // Step 2: Search the specified directory for an image file that matches the filename
            string[] imageExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif" };
            foreach (string extension in imageExtensions)
            {
                string imagePath = Path.Combine(imageDirectoryPath, fileNameWithoutExtension + extension);
                if (File.Exists(imagePath))
                {
                    // Step 3: Return the full UNC file path of the found image
                    return Path.GetFullPath(imagePath);
                }
            }

            // If no matching image file is found, return null or throw an exception
            return null;
        }

        private static List<string> ConvertFromJsonToLineList(JObject jsonObject)
        {
            var block = jsonObject["Value"]["Read"]["Blocks"][0];
            List<string> eCardStrings = new List<string>();

            foreach (var blockObject in block)
            {
                foreach (var line in blockObject)
                {
                    //var index = 0;

                    foreach (var lineObject in line)
                    {
                        var fieldValue = lineObject["Text"].ToString();
                        eCardStrings.Add(fieldValue);
                        //index++;
                    }

                }

            }
            return eCardStrings;
        }

        private static double getWHRatio(JObject jsonObject)
        {
            var cardBoundary = new Rectangle();
            var block = jsonObject["Value"]["Read"]["Blocks"][0];

            foreach (var blockObject in block)
            {
                foreach (var line in blockObject)
                {
                    var UpperLeftCoord = new Point();
                    var LowerRightCoord = new Point();
                    var upperLeftDone = false;

                    foreach (var lineObject in line)
                    {
                        //First, grab that upper left corner
                        if (!upperLeftDone)
                        {
                            UpperLeftCoord.X = int.Parse(lineObject["BoundingPolygon"][0]["X"].ToString());
                            UpperLeftCoord.Y = int.Parse(lineObject["BoundingPolygon"][0]["Y"].ToString());
                            cardBoundary.X = UpperLeftCoord.X;
                            cardBoundary.Y = UpperLeftCoord.Y;
                            upperLeftDone = true;
                        }
                        //Seek out the max low and max right edge 
                        int currentx = int.Parse(lineObject["BoundingPolygon"][0]["X"].ToString());
                        int currenty = int.Parse(lineObject["BoundingPolygon"][0]["Y"].ToString());

                        if (currentx > LowerRightCoord.X)
                            LowerRightCoord.X = currentx;

                        if (currenty > LowerRightCoord.Y)
                            LowerRightCoord.Y = currenty;
                    }

                    //now we have the max bounds of the text on the page
                    cardBoundary.Width = (LowerRightCoord.X - cardBoundary.X);
                    cardBoundary.Height = (LowerRightCoord.Y - cardBoundary.Y);
                }

            }

            return Math.Round((double)cardBoundary.Width / (double)cardBoundary.Height, 2);
        }

        private static JObject ReadJsonFile(string filePath)
        {
            using (StreamReader file = File.OpenText(filePath))
            {
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject jsonObject = (JObject)JToken.ReadFrom(reader);
                    return jsonObject;
                }
            }
        }

        private static string GetCsvFileName()
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
            string csvFileName = $"{timestamp}.csv";
            return csvFileName;
        }

        private static string GetFileNameWithoutExtension(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return string.Empty;
            }

            // Extract the filename with extension
            string fileNameWithExtension = Path.GetFileName(filePath);

            // Remove the extension
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileNameWithExtension);

            return fileNameWithoutExtension;
        }

        private static double CalculateDuplicatePercentage(JObject jObject)
        {
            List<string> stringList = new List<string>();

            var block = jObject["Value"]["Read"]["Blocks"][0];

            foreach (var blockObject in block)
            {
                foreach (var line in blockObject)
                {
                    foreach (var lineObject in line)
                    {
                        stringList.Add(lineObject["Text"].ToString());
                    }
                }

            }

            HashSet<string> uniqueStrings = new HashSet<string>();
            int duplicateCount = 0;

            foreach (string str in stringList)
            {
                if (!uniqueStrings.Add(str))
                {
                    duplicateCount++;
                }
            }

            double duplicatePercentage = (double)duplicateCount / (stringList.Count / 2) * 100;
            return Math.Round(duplicatePercentage,2);
        }

            private static eCardImageType DetectImageCertType(List<string> eCardStrings)
            {
            //double square card ACLS Provider
            if (eCardStrings[2].Contains("ACLS") &&
                eCardStrings[4].Contains("ACLS") &&
                eCardStrings[6].Contains("Provider") &&
                eCardStrings[8].Contains("Provider") &&
                eCardStrings.Count == 68)
                {
                    return eCardImageType.ACLS_PROVIDER_DOUBLE_SQUARE_68;
                }

            //test for acls_bls
            if (eCardStrings[1].Contains("ACLS") && 
                eCardStrings[2].Contains("Provider") && 
                eCardStrings.Count == 33)
                {
                    return eCardImageType.ACLS_PROVIDER_SINGLE_SQUARE_33;
                }

            if (eCardStrings[1].Contains("ACLS") &&
                eCardStrings[3].Contains("Instructor") &&
                eCardStrings.Count == 29)
            {
                return eCardImageType.ACLS_INSTRUCTOR_SINGLE_SQUARE_29;
            }

            if (eCardStrings[4].Contains("ACLS") &&
                eCardStrings[8].Contains("Instructor") &&
                eCardStrings.Count == 68)
            {
                return eCardImageType.ACLS_INSTRUCTOR_DOUBLE_SQUARE_INSTRUCTIONS_68;
            }
            
            if (eCardStrings[0].Contains("ADVANCED CARDIOVASCULAR LIFE SUPPORT") && 
                eCardStrings[1].Contains("ADVANCED CARDIOVASCULAR LIFE SUPPORT") && 
                eCardStrings[2].Contains("ACLS") &&
                eCardStrings[4].Contains("Instructor") &&
                eCardStrings.Count == 36)
            {
                return eCardImageType.ACLS_INSTRUCTOR_SINGLE_WALLET_INSTRUCTIONS_36;
            }

            if (eCardStrings[1].Contains("BLS") &&
                eCardStrings[3].Contains("Provider") &&
                eCardStrings.Count == 34)
            {
                return eCardImageType.BLS_SINGLE_SQUARE_34;
            }

            if (eCardStrings[2].Contains("BLS") &&
                eCardStrings[4].Contains("BLS") &&
                eCardStrings[1].Contains("BASIC LIFE SUPPORT") &&
                eCardStrings[6].Contains("Provider") &&
                eCardStrings.Count == 66)
            {
                return eCardImageType.BLS_DOUBLE_SQUARE_66;
            }

            if (eCardStrings[1].Contains("PALS") &&
                eCardStrings[5].Contains("Provider") &&
                eCardStrings.Count == 38)
            {
                return eCardImageType.PALS_SINGLE_SQUARE_38;
            }

            if (eCardStrings[2].Contains("PALS") &&
                eCardStrings[5].Contains("PALS") &&
                eCardStrings[8].Contains("Provider") &&
                eCardStrings[11].Contains("Provider") &&
                eCardStrings.Count == 70)
            {
                return eCardImageType.PALS_DOUBLE_SQUARE_70;
            }

            return eCardImageType.UNKNOWN;
            }
    }



    public enum eCardImageType
    {
        ACLS_PROVIDER_DOUBLE_SQUARE_68,
        ACLS_PROVIDER_SINGLE_SQUARE_33,
        ACLS_INSTRUCTOR_SINGLE_SQUARE_29,
        ACLS_INSTRUCTOR_DOUBLE_SQUARE_INSTRUCTIONS_68,
        ACLS_INSTRUCTOR_SINGLE_WALLET_INSTRUCTIONS_36,
        BLS_SINGLE_SQUARE_34,
        BLS_DOUBLE_SQUARE_66,
        PALS_SINGLE_SQUARE_38,
        PALS_DOUBLE_SQUARE_70,
        UNKNOWN
    }

    public class eCardCsvSheet
    {
        public string Filename { get; set; }
        public string LineCount { get; set; }
        public string WHRatio { get; set; }
        public string DupePct { get; set; }
        public string URL { get; set; }
        public string CertTitle { get; set; }
        public string LeftTitle { get; set; }
        public string FullName { get; set; }
        public string IssueDate { get; set; }
        public string RenewByDate { get; set; }
        public string TrainingCenterName { get; set; }
        public string InstructorName { get; set; }
        public string TrainingCenterId { get; set; }
        public string InstructorId { get; set; }
        public string TrainingCenterCityState { get; set; }
        public string EcardCode { get; set; }
        public string TrainingCenterPhoneNumber { get; set; }
        public string TrainingSiteName { get; set; }
    }
}
