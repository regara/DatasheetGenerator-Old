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
            string _aboveCodePerc =  DryPath(standard.Elements("EUseSummary"), "PctSavingsCmpTDV") + "%";
            string _spaceCool = Math.Round(Convert.ToDouble(standard.Elements("EnergyUse").SingleOrDefault(x => x.Element("Name").Value == "EU-SpcClg").Elements("PctImproveTDV").SingleOrDefault().Value), 1) + "%";


            /*********************** ATTIC ***********************/

            List<string> wallNames = new List<string>();
            List<string> wallMatArr = new List<string>();
            List<string> abvDeckArr = new List<string>();
            List<string> blwDeckArr = new List<string>();
            List<string> atticFloorArr = new List<string>();
            List<string> wallInsulArr = new List<string>();
            List<string> kneeWallArr = new List<string>();
            List<string> floorOvrGarArr = new List<string>();
            List<string> sidingOrStuccoArr = new List<string>();

            string _roofMatFormated;
            string _radientBarrier = "N/A";
            string _abvRoofDeck = "";
            string _blwRoofDeck = "";
            string _atticFloor = "";
            string _wallInsul24 = "-";
            string _wallInsul26 = "-";
            string _kneeWall = "-";
            string _floorOvrGar = "-";
            string _floorType = "";
            string _sidingOrStucco = "";



            if (userInput.Elements("Zone").Elements("SlabFloor").Count() > 0)
                _floorType = "Slab";
            if (userInput.Elements("Zone").Elements("FloorOverCrawl").Count() > 0)
            {
                if (_floorType != "")
                    _floorType += " / ";
                _floorType += "Raised Floor";
            }

            




            foreach (var atticFloor in userInput.Elements("Zone").Elements("CeilingBelowAttic"))
            {
                var temp = atticFloor.Element("Construction");
                if (!atticFloorArr.Contains(temp.Value.Split(' ')[0]) && temp.Value != "R-19 Attic Roof")
                    atticFloorArr.Add(temp.Value.Split(' ')[0]);
            }


            foreach (var ceil in wallNamePath)
            {
                if (!wallNames.Contains(Convert.ToString(ceil.Elements("Construction"))) &&
                   (string) ceil.Element("Construction") != "Attic Roof Garage")
                       wallNames.Add(Convert.ToString(ceil.Element("Construction").Value));
            }

            foreach (var roofingMat in wallNames)
            {
                var roofTypePath = roofMatPath.SingleOrDefault(x => x.Element("Name").Value == roofingMat).Element("RoofingLayer").Value;
                var radientBarrierPath = roofMatPath.SingleOrDefault(x => x.Element("Name").Value == roofingMat).Element("RadiantBarrier").Value;
                var aboveRoofDeckPath = roofMatPath.SingleOrDefault(x => x.Element("Name").Value == roofingMat).Element("AbvDeckInsulLayer").Value;
                var belowRoofDeckPath = roofMatPath.SingleOrDefault(x => x.Element("Name").Value == roofingMat).Element("CavityLayer").Value;


                switch (roofTypePath)
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

                if (radientBarrierPath == "1")
                    _radientBarrier = "Yes";

                if (!abvDeckArr.Contains(aboveRoofDeckPath) && aboveRoofDeckPath != "- no insulation -")
                    abvDeckArr.Add(aboveRoofDeckPath);



                if (!blwDeckArr.Contains(belowRoofDeckPath) && belowRoofDeckPath != "- no insulation -")
                    blwDeckArr.Add(belowRoofDeckPath);
            }

            foreach (var zone in userInput.Elements("Zone"))
            {
                foreach (var wall in zone.Elements("ExtWall"))
                {
                    if (!wallInsulArr.Contains(wall.Element("Construction").Value))
                    {
                        wallInsulArr.Add(wall.Element("Construction").Value);

                        if (Convert.ToInt32(wall.Element("Construction").Value.Split(' ')[0].Split('-')[1]) <= 15)
                        {
                            if (_wallInsul24.Length == 1)
                                _wallInsul24 = "";
                            if (_wallInsul24.Length > 1)
                                _wallInsul24 += " / ";

                            _wallInsul24 += wall.Element("Construction").Value.Split(' ')[0];
                            if (wall.Element("Construction").Value.Split(' ')[0].Split('-').Last().Contains("R"))
                                _wallInsul24 += wall.Element("Construction").Value.Split(' ')[0].Split('-').Last();
                        }
                        else
                        {
                            if (_wallInsul26.Length == 1)
                                _wallInsul26 = "";
                            if (_wallInsul26.Length > 1)
                                _wallInsul26 += " / ";

                            _wallInsul26 += wall.Element("Construction").Value.Split(' ')[0];
                            if (wall.Element("Construction").Value.Split(' ')[0].Split('-').Last().Contains("R"))
                                _wallInsul26 += wall.Element("Construction").Value.Split(' ')[0].Split('-').Last();
                        }
                    }

                    foreach (var wallI in zone.Elements("IntWall"))
                    {
                        if (!kneeWallArr.Contains(wallI.Element("Construction").Value.Split(' ')[0]))
                            kneeWallArr.Add(wallI.Element("Construction").Value.Split(' ')[0]);
                    }

                    foreach (var floor in zone.Elements("InteriorFloor"))
                    {
                        if (!floorOvrGarArr.Contains(floor.Element("Construction").Value.Split(' ')[0]))
                            floorOvrGarArr.Add(floor.Element("Construction").Value.Split(' ')[0]);
                    }
                }
            }

            foreach (var wallt in wallInsulArr)
            {
                var tempPath = roofMatPath.SingleOrDefault(x => x.Element("Name").Value == wallt)
                    .Element("WallExtFinishLayer").Value;
                if (!sidingOrStuccoArr.Contains(tempPath))
                {
                    sidingOrStuccoArr.Add(tempPath);


                    switch (tempPath)
                    {
                        case "Synthetic Stucco":
                            _sidingOrStucco += "Synthetic";
                            break;
                        case "3 Coat Stucco":
                            _sidingOrStucco += "3-Coat";
                            break;
                        case "Wood Siding/sheathing/decking":
                            _sidingOrStucco += "Siding/Sheathing";
                            break;
                        default:
                            _sidingOrStucco += "Unknown Type";
                            break;
                    }
                }

                
            }






                if (abvDeckArr.Count == 0)
                _abvRoofDeck = "-";
            else
                _abvRoofDeck = String.Join(" / ", abvDeckArr);

            if (blwDeckArr.Count == 0)
                _blwRoofDeck = "-";
            else
                _blwRoofDeck = String.Join(" / ", blwDeckArr);

            _roofMatFormated = String.Join(" / ", wallMatArr);
            _atticFloor = String.Join(" / ", atticFloorArr);
            _kneeWall = String.Join(" / ", kneeWallArr);
            _floorOvrGar = String.Join(" / ", floorOvrGarArr);



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



            /*********************** PLACEHOLDER ***********************/





            var property = (from p in doc.Descendants("Model").Where(n => (string)n.Attribute("Name") == "Proposed")
                            let proj = p.Element("Proj")
                            select new

                            {
                                photovoltaic = proj.Element("PVMinRatedPwrRpt").Value,
                                hERSIndex = "????",
                                planStucco = proj.Element("Name").Value,
                                fileName = proj.Element("ModelFile").Value,
                                squareFeet = proj.Element("CondFloorArea").Value,
                                stories = proj.Element("NumStories").Value,
                                glazing = Math.Round(Convert.ToDouble(proj.Element("CondWinAreaCFARat").Value),3) + "%",


   
                                reflectEmiss = proj.Elements("Attic").First().Element("RoofSolReflect").Value + " / " + proj.Elements("Attic").First().Element("RoofEmiss").Value,
                                kneeWall = "",
                                floorOverhang = "",
                                floorType = "",
                                seerEer = proj.Element("SCSysRpt").Element("MinCoolSEER").Value + " / " + proj.Element("SCSysRpt").Element("MinCoolEER").Value,
                                afue = "",
                                ductInsul = proj.Element("SCSysRpt").Element("MinDistribInsRval").Value,
                                wholeHouseFan = "",
                                fanWattage = "",
                                airflow = "",
                                ductTestingReq = "",
                                indoorAirQual = "",
                                refCharg = proj.Element("SCSysRpt").Element("HERSACCharg").Value,
                                seerVerif = proj.Element("SCSysRpt").Element("HERSSEER").Value,
                                eerVerif = proj.Element("SCSysRpt").Element("HERSEER").Value,
                                infiltration = "",
                                ductInConditioned = "",
                                lowLeakageAir = proj.Element("SCSysRpt").Element("LLAHUStatus").Value,
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
            datasheetCreator("PHOTO", (property.photovoltaic == "0") ? "N/A" : property.photovoltaic + " kWdc");
            datasheetCreator("HERS", property.hERSIndex);
            datasheetCreator("WALLTYPE", _sidingOrStucco);
            datasheetCreator("PLANNAME", _name);
            datasheetCreator("FILENAME", property.fileName);
            datasheetCreator("SQFT", property.squareFeet);
            datasheetCreator("ABVP", _aboveCodePerc);
            datasheetCreator("COOLP", _spaceCool);
            datasheetCreator("STORIES", property.stories);
            datasheetCreator("GLAZINGP", property.glazing.Split('.')[1].Insert(2, "."));
            datasheetCreator("ROOFMAT", _roofMatFormated);
            datasheetCreator("REFEM", property.reflectEmiss);
            datasheetCreator("ATTIC", _atticFloor);
            datasheetCreator("ABVRD", _abvRoofDeck);
            datasheetCreator("BLWRD", _blwRoofDeck);
            datasheetCreator("RADIENT", _radientBarrier);
            datasheetCreator("WALL24", _wallInsul24);
            datasheetCreator("WALL26", _wallInsul26);
            datasheetCreator("KNEEWALL", _kneeWall);
            datasheetCreator("OVERG", _floorOvrGar);
            datasheetCreator("FLOORTYPE", _floorType);
            datasheetCreator("SEEREER", property.seerEer);
            datasheetCreator("AFUE", "Placeholder");
            datasheetCreator("DUCTINS", "R-" + property.ductInsul);
            datasheetCreator("WHF", "Placeholder");
            datasheetCreator("FANWAT", "Placeholder");
            datasheetCreator("AIRFLOW", "Placeholder");
            datasheetCreator("DUCTTEST", "Placeholder");
            datasheetCreator("CFM", "Placeholder");
            datasheetCreator("REFCHARGE", (property.refCharg == "1") ? "Yes" : "-");
            datasheetCreator("SEERVERIF", (property.seerVerif == "1") ? "Yes" : "-");
            datasheetCreator("EERVERIF", (property.eerVerif == "1") ? "Yes" : "-");
            datasheetCreator("INFILTRATION", "Placeholder");
            datasheetCreator("DUCTCOND", "Placeholder");
            datasheetCreator("LOWLEAK", (property.lowLeakageAir == "Has Low Leakage Air Handler") ? "Yes" : "-");
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
