using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Win32;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {

            /*********************** btn on click upload XML ***********************/

            OpenFileDialog CBECCXML = new OpenFileDialog();
            CBECCXML.Filter = "XML Files (*.xml)|*.xml";
            CBECCXML.FilterIndex = 1;

            CBECCXML.Multiselect = true;
            CBECCXML.ShowDialog();

            string s = CBECCXML.FileName;
            XDocument doc = XDocument.Load(s);


            /*********************** DRY PATH METHOD ***********************/

            string DryPath(IEnumerable<XElement> path, XName element)
            {
                return Convert.ToString(path.Elements(element).First().Value);
            }

            /*********************** XMLNODE - PATHS ***********************/

            var userInput = doc.Descendants("Model").Take(1).Elements("Proj");
            var proposed = doc.Descendants("Model").Skip(1).Take(1).Elements("Proj");
            var standard = doc.Descendants("Model").Skip(2).Take(1).Elements("Proj");

            //            var wallPath = userInput.Elements("Zone").Elements("ExtWall");
            //            var wallPath = userInput.Elements("Zone").Elements("ExtWall");
            var windowPath = proposed.Descendants("Win");
            var wallNamePath = userInput.Descendants("Attic");
            var roofMatPath = doc.Descendants("Model").Take(1).Descendants("Cons");

            /*********************** Misc General Data ***********************/

            string[] _softversionTemp = Convert.ToString(userInput.Elements("SoftwareVersion").First()).Split(' ')[1].Split('.');
            string _softversion = "CBECC V" + _softversionTemp[1] +"."+ _softversionTemp[2];


            string _name = DryPath(userInput, "Name");
            string _city = DryPath(userInput, "City");
            string _climateZone = DryPath(userInput, "ClimateZone").Split(' ')[0].Substring(2,2);
            string _aboveCodePerc =  DryPath(standard.Elements("EUseSummary"), "PctSavingsCmpTDV");
            //            string _coolingImprove = DryPath(standard, );









            /*********************** ATTIC ***********************/

            //            roofing = proj.Elements("Cons")[].Elements("RoofingLayer"),

            //          var proposed 

            List<string> wallNames = new List<string>();
            List<string> wallMatArr = new List<string>();
            string _roofMatFormated;



            foreach (var ceil in wallNamePath)
            {
                if (!wallNames.Contains(Convert.ToString(ceil.Elements("Construction"))) &&
                   (string) ceil.Element("Construction") != "Attic Roof Garage")
                       wallNames.Add(Convert.ToString(ceil.Element("Construction").Value));
            }

            foreach (var roofingMat in wallNames)
            {
                var tempRoofLayer = roofMatPath.SingleOrDefault(x => x.Element("Name").Value == roofingMat).Element("RoofingLayer").Value;

                switch (tempRoofLayer)
                {
                    case "Light Roof (Asphalt Shingle)":
                        wallMatArr.Add("Asphalt");
                        break;
                    case "5 PSF (Normal Gravel)":
                        wallMatArr.Add("Built Up");
                        break;
                    case "10 PSF (RoofTile)":
                        wallMatArr.Add("Tile");
                        break;
                    case "15 PSF (Heavy Ballast or Pavers)":
                        wallMatArr.Add("Heavy Ballast");
                        break;
                    case "25 PSF (Very Heavy Ballast or Pavers)":
                        wallMatArr.Add("Very Heavy Ballast");
                        break;
                    case "Light Roof (Metal Tile)":
                        wallMatArr.Add("Metal");
                        break;
                    default:
                        wallMatArr.Add("??????");
                        break;

                }                
            }
            _roofMatFormated = String.Join(" / ", wallMatArr);

            /*********************** WINDOWS ***********************/

            List<string> windowsUV = new List<string>();
            List<string> windowsSHGC = new List<string>();
            List<string> windowTypes = new List<string>();

            foreach (var windowObj in windowPath)
            {
                if (!windowTypes.Contains(Convert.ToString(windowObj.Element("WinType"))))
                {
                    windowTypes.Add(Convert.ToString(windowObj.Element("WinType")));
                    windowsUV.Add(Convert.ToString(windowObj.Element("NFRCUfactor")));
                    windowsSHGC.Add(Convert.ToString(windowObj.Element("NFRCSHGC")));
                }
            };

            //            foreach (string wind in windowTypes)
            //            {
            //                Console.WriteLine(wind);
            //            }
            //            foreach (string wind in windowsSHGC)
            //            {
            //                Console.WriteLine(wind);
            //            }
            //            foreach (string wind in windowsUV)
            //            {
            //                Console.WriteLine(wind);
            //            }


            /*********************** Wall ***********************/
            List<string> wallValues = new List<string>();
            List<string> wallType = new List<string>();

//            foreach (var extwall in wallPath)
//            {
//                if (!wallValues.Contains(Convert.ToString(extwall.Element("Construction"))))
//
//                    wallValues.Add(Convert.ToString(extwall.Element("Construction")));
//            }
//
//            foreach (string wallFinish in wallType)
//            {
//                if (!wallType.Contains(Convert.ToString(extwall.Element("Construction"))))
//                {
//                    //    LEFT OFF HERE - 4Each val in array check every cons.child(name.value = arrVal)                doc.Descendants("Model").Take(1).Descendants("Cons");
//                }
//
//            }



            foreach (string wallt in wallValues)
            {
                Console.WriteLine(wallt);
                Console.WriteLine(wallt);

            }

            foreach (string wallt in wallType)
            {
                Console.WriteLine(wallt);
                Console.WriteLine(wallt);

            }

            //            wallCoatType = "IN THE CONSTRUCT FOR EACH WALL",



            /*********************** PLACEHOLDER ***********************/





            var property = (from p in doc.Descendants("Model").Where(n => (string)n.Attribute("Name") == "Proposed")
                            let proj = p.Element("Proj")
                            select new

                            {
                                photovoltaic = proj.Element("PVMinRatedPwrRpt").Value + " kWDC",
                                hERSIndex = "????",
                                planStucco = proj.Element("Name").Value,
                                fileName = proj.Element("ModelFile").Value,
                                squareFeet = proj.Element("CondFloorArea").Value,
                                stories = proj.Element("NumStories").Value,
                                glazing = proj.Element("CondWinAreaCFARat").Value + "%",



                                reflectEmiss = proj.Elements("Attic").First().Element("RoofSolReflect").Value + " / " + proj.Elements("Attic").First().Element("RoofEmiss").Value,
                                roofingVal = "",
                                //                  Leave out unless can append xml appropriatly        atticAboveD = proj.Elements("Attic").First().Element("AboveDeckRoofIns").Value,
                                atticBelowD = proj.Elements("Attic").First().Element("BelowDeckRoofIns").Value,
                                radientBar = "",
                                kneeWall = "",
                                floorOverhang = "",
                                floorType = "",
                                seerEer = proj.Element("SCSysRpt").Element("MinCoolSEER").Value + " / " + proj.Element("SCSysRpt").Element("MinCoolEER").Value,
                                afue = "",
                                ductInsul = "",
                                wholeHouseFan = "",
                                fanWattage = "",
                                airflow = "",
                                ductTestingReq = "",
                                indoorAirQual = "",
                                refCharg = "",
                                seerVerif = "", //zone.HVACSysAssigned
                                eerVerif = "",
                                infiltration = "",
                                ductInConditioned = "",
                                lowLeakageAir = "",
                                burriedDuct = "",
                                surfaceArea = "",
                                insulInspect = "",
                                fuelType = proj.Element("GasType").Value,

                                //      Get value of each proj.<DHWSys> as WH Then search for proj.<DHWHeater> with a child <name>.value = WH
                                //                    waterHeater = proj.Element("DHWHeater").Element("EnergyFactor").Value,

                                distribution = "Standard"
                            }).SingleOrDefault();






            List<string> result = Convert.ToString(property).Split(',').ToList();

            foreach (var prop in result)
            {
                //                            var props = prop.Split('='); ***** if theres an extra row, it will be hard to tell visually if you remove the key
                //                            Console.WriteLine(props[1]);

                //                            Console.WriteLine(prop);

            }


































//            XDocument xmlFile = XDocument.Load("XMLTemplates/Test.xml");
//            var query = from c in xmlFile.Elements("catalog").Elements("book")
//                        select c;
//            foreach (XElement book in query)
//            {
//                book.Attribute("attr1").Value = "MyNewValue";
//            }
//            xmlFile.Save("books.xml");
//
//
//
//
//
////            var nodec = xmlDocument.SelectSingleNode("//*[@id='10001']");
////            Console.WriteLine(nodec.InnerText);
//            
//
//

            XmlDocument datasheet = new XmlDocument();
            datasheet.Load(@"..\..\XMLTemplates\2016 Datasheet - 1 Column.xml");


            void datasheetCreator(string ID, string textToAppend)
            {
                datasheet.SelectSingleNode("//*[@id='"+ ID + "']").InnerText = textToAppend;
            }


            datasheetCreator("TODAYDATE", DateTime.Now.ToString("MM/dd/yyyy"));
            datasheetCreator("SOFTVERSION", _softversion);
            datasheetCreator("PROJECTNAME", _name);
            datasheetCreator("CITY", _city);
            datasheetCreator("CLIMATEZONE", _climateZone);
            datasheetCreator("PHOTO", property.photovoltaic);
            datasheetCreator("HERS", property.hERSIndex);
            datasheetCreator("PLANNAME", _name);
            datasheetCreator("FILENAME", property.fileName);
            datasheetCreator("SQFT", property.squareFeet);
            datasheetCreator("ABVP", _aboveCodePerc);
            datasheetCreator("COOLP", "Placeholder");
            datasheetCreator("STORIES", property.stories);
            datasheetCreator("GLAZINGP", property.glazing);
            datasheetCreator("ROOFMAT", _roofMatFormated);
            datasheetCreator("REFEM", property.reflectEmiss);
            datasheetCreator("ATTIC", "Placeholder");
            datasheetCreator("ABVRD", "Placeholder");
            datasheetCreator("BLWRD", property.atticBelowD);
            datasheetCreator("RADIENT", "Placeholder");
            datasheetCreator("WALL24", "Placeholder");
            datasheetCreator("WALL26", "Placeholder");
            datasheetCreator("KNEEWALL", "Placeholder");
            datasheetCreator("OVERG", "Placeholder");
            datasheetCreator("FLOORTYPE", "Placeholder");
            datasheetCreator("SEEREER", property.seerEer);
            datasheetCreator("AFUE", "Placeholder");
            datasheetCreator("DUCTINS", "Placeholder");
            datasheetCreator("WHF", "Placeholder");
            datasheetCreator("FANWAT", "Placeholder");
            datasheetCreator("AIRFLOW", "Placeholder");
            datasheetCreator("DUCTTEST", "Placeholder");
            datasheetCreator("CFM", "Placeholder");
            datasheetCreator("REFCHARGE", "Placeholder");
            datasheetCreator("SEERVERIF", "Placeholder");
            datasheetCreator("EERVERIF", "Placeholder");
            datasheetCreator("INFILTRATION", "Placeholder");
            datasheetCreator("DUCTCOND", "Placeholder");
            datasheetCreator("LOWLEAK", "Placeholder");
            datasheetCreator("BURRIEDDUCT", "Placeholder");
            datasheetCreator("SURFAREA", "Placeholder");
            datasheetCreator("INSULINSPECT", "Placeholder");
            datasheetCreator("FUELTYPE", property.fuelType);
            datasheetCreator("UEF", "Placeholder");
            datasheetCreator("DISTRIBUTION", property.distribution);


















            datasheetCreator("WIN1", "Placeholder");
            datasheetCreator("WIN2", "Placeholder");
            datasheetCreator("WIN3", "Placeholder");
            datasheetCreator("WIN4", "Placeholder");

            datasheetCreator("UVAL1", "Placeholder");
            datasheetCreator("UVAL2", "Placeholder");
            datasheetCreator("UVAL3", "Placeholder");
            datasheetCreator("UVAL4", "Placeholder");

            datasheetCreator("SHGC1", "Placeholder");
            datasheetCreator("SHGC2", "Placeholder");
            datasheetCreator("SHGC3", "Placeholder");
            datasheetCreator("SHGC4", "Placeholder");






            datasheet.Save(@"..\..\Testttttttt.doc");



        }
    }
}
