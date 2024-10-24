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
            int rowNum = 0;
            //int typeNotFound = 0;
            int success = 0;
            int notRecognized = 0;

            foreach (string jsonFilePath in Directory.EnumerateFiles(jsonDirectoryPath, "*.json"))
            {
                try
                {
                    rowNum++;

                    JObject jsonObject = ReadJsonFile(jsonFilePath);
                    List<CardTextComponent> componentList = ConvertFromJsonToTextCardObjectList(jsonObject);

                    List<string> linesOnTheCard = ConvertFromJsonToLineList(jsonObject);

                    var labelList = RetrieveLabels(linesOnTheCard);

                    if(!CardSanityCheck(linesOnTheCard, labelList))
                    {
                        Console.WriteLine($"#{rowNum}-Failure-Fields not recognized {GetFileNameWithoutExtension(jsonFilePath)}");
                        notRecognized++;
                        continue;
                    }

                    eCardCsvSheet csvSheet = FillOutCSVSheet(labelList, componentList, jsonFilePath);

                    eCardCSVList.Add(csvSheet);

                    Console.WriteLine($"#{rowNum}-Success {GetFileNameWithoutExtension(jsonFilePath)}");
                    success++;

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"#{rowNum}-Exception {GetFileNameWithoutExtension(jsonFilePath)}, {ex.Message}");
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

            Console.WriteLine($"Complete at {DateTime.Now} Success={success} NotRecognized={notRecognized}");
            //Console.WriteLine($"RQI-{rqiCount}");
            Console.ReadKey();
        }

        private static bool ListContainsSubstring(List<string> list, string subString)
        {
            if (list.FirstOrDefault(s => s.Contains(subString)) != null)
            {
                return true;
            }
            return false;
        }

        private static bool CardSanityCheck(List<string> linesOnTheCard, LabelList labelList)
        {       if (labelList is null)
                    return false;
                if(linesOnTheCard.Count == 0)
                    return false;
                if( !linesOnTheCard.Contains(labelList.IssueDate))
                    return false;
                if (!linesOnTheCard.Contains(labelList.RenewBy))
                    return false;
                if (!linesOnTheCard.Contains(labelList.ECardCode))
                    return false;
                if(!ListContainsSubstring(linesOnTheCard, labelList.Name))
                    return false;

            return true;
        }

        private static eCardCsvSheet FillOutCSVSheet(LabelList labelList, List<CardTextComponent> componentList, string filePath)
        {
            var eCardSheet = new eCardCsvSheet();

            eCardSheet.IssueDate = FindDataByLabel(componentList, labelList.IssueDate);
            eCardSheet.RenewByDate = FindDataByLabel(componentList, labelList.RenewBy);
            eCardSheet.EcardCode = FindDataByLabel(componentList, labelList.ECardCode);
            eCardSheet.FullName = FindDataByLabel(componentList, labelList.Name, SearchDirection.UP);
            eCardSheet.Filename = filePath;
            eCardSheet.CertTitle = componentList[0].Text;

            return eCardSheet;
        }

        private static LabelList RetrieveLabels(List<string> fieldList)
        {            
            //The standard template
            if (ListContainsSubstring(fieldList, "has successfully completed the cognitive") &&
                fieldList.Contains("Issue Date") &&
                fieldList.Contains("Renew By") &&
                fieldList.Contains("eCard Code"))
            {
                return new LabelList
                {
                    IssueDate = "Issue Date",
                    RenewBy = "Renew By",
                    ECardCode = "eCard Code",
                    Name = "has successfully completed the cognitive and"
                };
            }

            //The standard template with a spaced e Card code
            if (ListContainsSubstring(fieldList, "has successfully completed the cognitive") &&
                fieldList.Contains("Issue Date") &&
                fieldList.Contains("Renew By") &&
                fieldList.Contains("e Card Code"))
            {
                return new LabelList
                {
                    IssueDate = "Issue Date",
                    RenewBy = "Renew By",
                    ECardCode = "e Card Code",
                    Name = "has successfully completed the cognitive and"
                };
            }

            //The Gold Stamp RQI template
            if (ListContainsSubstring(fieldList, "This is to verify that") &&
                ListContainsSubstring(fieldList, "has demonstrated competence in") &&
                fieldList.Contains("Date of last activity:") &&
                fieldList.Contains("eCredential valid until:") &&
                fieldList.Contains("eCredential number:"))
            {
                return new LabelList
                {
                    IssueDate = "Date of last activity:",
                    RenewBy = "eCredential valid until:",
                    ECardCode = "eCredential number:",
                    Name = "has demonstrated competence in "
                };
            }

            return null;
        }

        private static List<CardTextComponent> ConvertFromJsonToTextCardObjectList(JObject jsonObject)
        {
            var block = jsonObject["Value"]["Read"]["Blocks"][0];
            List<CardTextComponent> textComponents = new List<CardTextComponent>();

            foreach (var blockObject in block)
            {
                foreach (var line in blockObject)
                {
                    foreach (var lineObject in line)
                    {
                        var fieldValue = lineObject["Text"].ToString();
                        var boundingBox = lineObject["BoundingPolygon"];
                        var pointUL = boundingBox[0];
                        var pointUR = boundingBox[1];
                        var pointLL = boundingBox[2];
                        var pointLR = boundingBox[3];
                        var textComponent = new CardTextComponent(
                            fieldValue,
                            int.Parse(pointUL["X"].ToString()),
                            int.Parse(pointUL["Y"].ToString()),
                            int.Parse(pointUR["X"].ToString()),
                            int.Parse(pointUR["Y"].ToString()),
                            int.Parse(pointLL["X"].ToString()),
                            int.Parse(pointLL["Y"].ToString()),
                            int.Parse(pointLR["X"].ToString()),
                            int.Parse(pointLR["Y"].ToString())
                            );
                        textComponents.Add(textComponent);
                    }

                }

            }

            return textComponents;
        }

 

        private static string FindDataByLabel(
            List<CardTextComponent> componentList, 
            String labelName, 
            SearchDirection searchDirection=SearchDirection.DOWN)
        {
            String foundStr = string.Empty;

            //validate params
            if (componentList is null || componentList.Count ==0 || string.IsNullOrEmpty(labelName))
            {
                Console.WriteLine($"FindDataByLabel has bad params");
                return foundStr;
            }    

            //find the label first 
            System.Drawing.Point labelGeometricCenter = Point.Empty;

            CardTextComponent labelComponent = null;

            foreach (var component in componentList)
            {
                if (component.Text.Contains(labelName))
                {
                    labelComponent = component;
                    break;
                }
            }

            if (labelComponent is null)
            {
                Console.WriteLine($"Label {labelName} not found");
                return foundStr;
            }

            CardTextComponent closestComponent = null;
            int closestDistance = int.MaxValue;

            for (int i = 1; i < componentList.Count; i++)
            {
                CardTextComponent nextComponent = componentList[i];

                //dont even consider data above a label, and dont consider the label itself.

                if (searchDirection == SearchDirection.DOWN)
                {
                    if (nextComponent.GeometricCenter().Y - (int)(nextComponent.TextRectangle.Height / 2) <= labelComponent.GeometricCenter().Y ||
                        nextComponent.Text.Contains(labelName))
                    {
                        continue;
                    }
                }
                else //search UP
                {
                    if (nextComponent.GeometricCenter().Y + (int)(nextComponent.TextRectangle.Height / 2) >= labelComponent.GeometricCenter().Y ||
                        nextComponent.Text.Contains(labelName))
                    {
                        continue;
                    }
                }





                //calculate distance from label
                int distanceFromLabel = CalculateDistance(labelComponent.GeometricCenter(), nextComponent.GeometricCenter());

                if (distanceFromLabel < closestDistance)
                {
                    closestDistance = distanceFromLabel;
                    closestComponent = nextComponent;
                }

            }

            if (closestComponent != null)
            {
                foundStr = closestComponent.Text;
            }

            return foundStr;

        }

        private static int CalculateDistance(Point point1, Point point2)
        {
            int deltaX = point1.X - point2.X;
            int deltaY = point1.Y - point2.Y;
            return (int)Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
        }

        class LabelList
        {
            public string IssueDate { get; set; }
            public string RenewBy { get; set; }
            public string ECardCode { get; set; }
            public string Name { get; set; }
        }

        class CardTextComponent
        {
            public string Text { get; set; }
            public Rectangle TextRectangle { get; set; }

            public CardTextComponent(string text, int upperLeftX, int upperLeftY, int upperRightX, int upperRightY, int lowerRightX, int lowerRightY, int lowerLeftX, int lowerLeftY)
            {
                Text = text;
                int minX = Math.Min(Math.Min(upperLeftX, upperRightX), Math.Min(lowerRightX, lowerLeftX));
                int minY = Math.Min(Math.Min(upperLeftY, upperRightY), Math.Min(lowerRightY, lowerLeftY));
                int maxX = Math.Max(Math.Max(upperLeftX, upperRightX), Math.Max(lowerRightX, lowerLeftX));
                int maxY = Math.Max(Math.Max(upperLeftY, upperRightY), Math.Max(lowerRightY, lowerLeftY));
                TextRectangle = new Rectangle(minX, minY, maxX - minX, maxY - minY);
            }

            public Point GeometricCenter()
            {
                int centerX = TextRectangle.Left + (TextRectangle.Width / 2);
                int centerY = TextRectangle.Top + (TextRectangle.Height / 2);
                return new Point(centerX, centerY);
            }
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

        
    }

    public enum SearchDirection
    {
        UP,
        DOWN
    }

    public enum TemplateType
    {
        STANDARD_ECARD,
        RQI,
        UNKNOWN
    }

    public class eCardCsvSheet
    {
        public string CertTitle { get; set; }
        public string FullName { get; set; }
        public string IssueDate { get; set; }
        public string RenewByDate { get; set; }
        public string EcardCode { get; set; }
        public string Filename { get; set; }
    }
}
