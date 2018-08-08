using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

            OpenFileDialog cbeccxml = new OpenFileDialog
            {
                Filter = "XML Files (*.xml)|*.xml",
                FilterIndex = 1,
                Multiselect = true
            };

            cbeccxml.ShowDialog();

            string s = cbeccxml.FileName;
            XDocument doc = XDocument.Load(s);

            /*********************** DRY PATH METHOD ***********************/

            string DryPath(IEnumerable<XElement> path, XName element)
            {
                return Convert.ToString(path.Elements(element).FirstOrDefault()?.Value);
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

            string[] _softversionTemp = Convert.ToString(userInput.Elements("SoftwareVersion").First().Value).Split(' ')[1].Split('.');
            string _softversion = "CBECC V" + _softversionTemp[1] +"."+ _softversionTemp[2];

            string _name = DryPath(userInput, "Name");
            string _city = DryPath(userInput, "City");

            string _climateZone = DryPath(userInput, "ClimateZone").Split(' ')[0].Split('Z')[1];
            //            string _aboveCodePerc =  DryPath(standard.Elements("EUseSummary"), "PctSavingsCmpTDV");
            string _aboveCodePerc = standard.Elements("EUseSummary").Elements("PctSavingsCmpTDV").FirstOrDefault()?.Value;
            string _spaceCool = Math.Round(Convert.ToDouble(standard.Elements("EnergyUse").SingleOrDefault(x => x.Element("Name").Value == "EU-SpcClg")?.Elements("PctImproveTDV").SingleOrDefault().Value), 1) + "%";

            Console.WriteLine("abv ----" + _aboveCodePerc);
            foreach (var VARIABLE in standard.Elements("EUseSummary"))
            {
                Console.WriteLine("Var + " + VARIABLE);
            }
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
            List<string> waterHeaterArr = new List<string>();


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
            string _indoorAirQual = doc.Descendants("Model").Skip(1).Take(1).Elements("IAQVentRpt").SingleOrDefault()?.Element("IAQCFM")?.Value;
            string _insulConsQual = userInput.Elements("InsulConsQuality").SingleOrDefault()?.Value;
            string _airLeakage = (userInput.Elements("ACH50").SingleOrDefault()?.Value.Length == 1) ? userInput.Elements("ACH50").SingleOrDefault()?.Value +".00%" : userInput.Elements("ACH50").SingleOrDefault()?.Value + "%";
            string _buriedDuct = "-";
            string _surfaceArea = "-";
            string _ductInConditioned = "-";
            string _waterHeater = "";
            string _Qii = doc.Descendants("Model").Skip(1).Take(1).Single()?.Element("HERSOther")?.Element("QII")?.Value;

            if (userInput.Elements("Zone").Elements("SlabFloor").Count() > 0)
                _floorType = "Slab";
            if (userInput.Elements("Zone").Elements("FloorOverCrawl").Count() > 0)
            {
                if (_floorType != "")
                    _floorType += " / ";
                _floorType += "Raised Floor";
            }



            foreach (var buriedDucts in doc.Descendants("Model").Take(1).Elements("HVACSys"))
            {

                var temp = buriedDucts.Element("DistribSystem")?.Value;
                var query = doc.Descendants("Model").Take(1)?.Elements("HVACDist").SingleOrDefault(x => x.Element("Name").Value == temp);

                if (query?.Element("AreBuried")?.Value == "1")
                {
                    _buriedDuct = "Yes";
                    break;
                }

                if (query?.Element("DuctDesign")?.Value == "1")
                {
                    _surfaceArea = "Yes";
                    break;
                }

                if (query?.Element("_ductInConditioned")?.Value == "Ducts located within the conditioned space (except < 12 lineal ft)" ||
                    query?.Element("_ductInConditioned")?.Value == "Ducts located entirely in conditioned space" ||
                    query?.Element("_ductInConditioned")?.Value == "Verified low-leakage ducts entirely in conditioned space")
                {
                    _ductInConditioned = "Yes";
                    break;
                }
            }

            foreach (var waterHeater in doc.Descendants("Model").Take(1).Elements("DHWSys"))
            {
                var temp = waterHeater.Element("DHWHeater")?.Value;
                var temp2 = doc.Descendants("Model").Take(1).Elements("DHWHeater").SingleOrDefault(x => x.Element("Name")?.Value == temp);

                if (!waterHeaterArr.Contains(temp))
                {
                    if (waterHeaterArr.Count > 0)
                        _waterHeater += " + ";

                    waterHeaterArr.Add(temp);
                    if (temp2.Element("EnergyFactor")?.Value.Split('.').ToList().ElementAt(1)?.Length == 1)
                        temp2.Element("EnergyFactor").Value += "0";
                    _waterHeater += temp2.Element("EnergyFactor")?.Value + "(" + temp2.Element("TankVolume")?.Value +") ";
                }
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
                       wallNames.Add(Convert.ToString(ceil.Element("Construction")?.Value));
            }

            foreach (var roofingMat in wallNames)
            {
                var roofTypePath = roofMatPath.SingleOrDefault(x => x.Element("Name").Value == roofingMat)?.Element("RoofingLayer")?.Value;
                var radientBarrierPath = roofMatPath.SingleOrDefault(x => x.Element("Name").Value == roofingMat)?.Element("RadiantBarrier")?.Value;
                var aboveRoofDeckPath = roofMatPath.SingleOrDefault(x => x.Element("Name").Value == roofingMat)?.Element("AbvDeckInsulLayer")?.Value;
                var belowRoofDeckPath = roofMatPath.SingleOrDefault(x => x.Element("Name").Value == roofingMat)?.Element("CavityLayer" )?.Value;


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
                    blwDeckArr.Add(belowRoofDeckPath.Replace(" ", "-"));
            }

            foreach (var zone in userInput.Elements("Zone"))
            {
                foreach (var wall in zone.Elements("ExtWall"))
                {
                    if (!wallInsulArr.Contains(wall.Element("Construction").Value))
                    {
                        Console.WriteLine(wall.Element("Construction").Value);
                        wallInsulArr.Add(wall.Element("Construction").Value);

                        void WallFormater(string wallPathVal)
                        {
                            string firstWordInString = wallPathVal.Split(' ')[0];
                            string rVal = Regex.Split(firstWordInString, @"\D+")[1];

                            if (firstWordInString == "existing")
                            {
                                //ignore?
                            } else if (Convert.ToInt32(rVal) <= 15)
                            {

                                if (_wallInsul24.Length == 1)
                                    _wallInsul24 = "";

                                if (!_wallInsul24.Contains(firstWordInString))
                                {
                                    if (_wallInsul24.Length > 1)
                                       _wallInsul24 += " / ";

                                    _wallInsul24 += firstWordInString;
                                    if (wall.Element("Construction").Value.Split(' ').Last().Contains("R"))
                                    _wallInsul24 += " + " + wall.Element("Construction").Value.Split(' ').Last().Split('-')[1];
                                }
                            }
                            else
                            {
                                if (_wallInsul26.Length == 1)
                                   _wallInsul26 = "";

                                if (!_wallInsul26.Contains(firstWordInString))
                                {
                                    if (_wallInsul26.Length > 1)
                                       _wallInsul26 += " / ";

                                    _wallInsul26 += firstWordInString;
                                    if (wall.Element("Construction").Value.Split(' ').Last().Contains("R"))
                                        _wallInsul26 += " + " + wall.Element("Construction").Value.Split(' ').Last().Split('-')[1];
                                }
                            }
                        }
                        WallFormater(wall.Element("Construction").Value);
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
                var tempPath = roofMatPath.SingleOrDefault(x => x.Element("Name")?.Value == wallt)
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

            _wallInsul24 = string.Join(" / ", _wallInsul24.Split('/').Distinct().OrderBy(x => x));
            _wallInsul26 = string.Join(" / ", _wallInsul26.Split('/').Distinct().OrderBy(x => x));
            _kneeWall = string.Join(" / ", _kneeWall.Split('/').Distinct().OrderBy(x => x));


            /*********************** WINDOWS ***********************/

            List<string> windowsUV = new List<string>();
            List<string> windowsSHGC = new List<string>();
            List<string> windowTypes = new List<string>();

            foreach (var zones in proposed.Elements("Zone"))
            {
                foreach (var walls in zones.Elements("ExtWall"))
                {
                    
                    foreach (var windows in walls.Elements("Win"))
                    {
                        if (!windowsUV.Contains(windows.Element("NFRCUfactor")?.Value) ||
                            !windowsSHGC.Contains(windows.Element("NFRCSHGC")?.Value) ||
                            !windowTypes.Contains(windows.Element("WinType")?.Value))
                        {

                            windowsUV.Add(windows.Element("NFRCUfactor")?.Value);
                            windowsSHGC.Add(windows.Element("NFRCSHGC")?.Value);
                            windowTypes.Add(windows.Element("WinType")?.Value);

                        }
                    }
                }
            }

            /*********************** PLACEHOLDER ***********************/

            string _wholeHouseFan = "";

            foreach (var whf in proposed.Elements("UnitClVentCFM"))
            {
                if (_wholeHouseFan.Length != 0 && whf.Value != "0")
                    _wholeHouseFan += " + ";
                if (whf.Value != "0")
                    _wholeHouseFan += whf.Value;
            }

            _wholeHouseFan = string.Join("+", _wholeHouseFan.Split('+').Distinct().OrderByDescending(x => x));

            _wholeHouseFan = "Yes ( " + _wholeHouseFan + ")";
            if (_wholeHouseFan == "Yes ( )")
                _wholeHouseFan = "-";

            var property = (from p in doc.Descendants("Model").Where(n => (string)n.Attribute("Name") == "Proposed")
                            let proj = p.Element("Proj")
                            select new
                            {
                                photovoltaic = proj.Element("PVMinRatedPwrRpt").Value,
                                planStucco = proj.Element("Name").Value,
                                fileName = proj.Element("ModelFile").Value,
                                squareFeet = proj.Element("CondFloorArea").Value,
                                stories = proj.Element("NumStories").Value,
                                glazing = Math.Round(Convert.ToDouble(proj.Element("CondWinAreaCFARat").Value), 3),



                                reflectEmiss = proj.Elements("Attic")?.FirstOrDefault()?.Element("RoofSolReflect").Value + " / " + proj.Elements("Attic")?.FirstOrDefault()?.Element("RoofEmiss").Value,
                                kneeWall = "",
                                floorOverhang = "",
                                floorType = "",
                                seerEer = proj.Element("SCSysRpt")?.Element("MinCoolSEER")?.Value + " / " + proj.Element("SCSysRpt")?.Element("MinCoolEER")?.Value,
                                afue = proj.Element("SCSysRpt")?.Element("MinHeatEffic")?.Value,
                                ductInsul = proj.Element("SCSysRpt")?.Element("MinDistribInsRval")?.Value,
                                fanWattage = proj.Element("SCSysRpt")?.Element("HERSFanEff")?.Value,
                                airflow = proj.Element("SCSysRpt")?.Element("HERSAHUAirFlow")?.Value,
                                airflowVal = proj.Element("SCSysRpt")?.Element("MinCoolCFMperTon")?.Value,
                                ductTestingReq = proj.Element("SCSysRpt")?.Element("HERSDuctLeakage")?.Value,
                                ductTestingVal = proj.Element("SCSysRpt")?.Element("HERSDuctLkgRptMsg")?.Value,
                                indoorAirQual = "",
                                refCharg = proj.Element("SCSysRpt")?.Element("HERSACCharg")?.Value,
                                seerVerif = proj.Element("SCSysRpt")?.Element("HERSSEER")?.Value,
                                eerVerif = proj.Element("SCSysRpt")?.Element("HERSEER")?.Value,
                                infiltration = "",
                                ductInConditioned = "",
                                lowLeakageAir = proj.Element("SCSysRpt")?.Element("LLAHUStatus")?.Value,
                                insulInspect = "",
                                fuelType = proj.Element("GasType")?.Value,

                                distribution = "Standard"
                            }).SingleOrDefault();

            List<string> result = Convert.ToString(property).Split(',').ToList();

            XmlDocument datasheet = new XmlDocument();

          datasheet.Load(@"..\..\XMLTemplates\2016 Datasheet - 1 Column.xml");
//            datasheet.Load(@"XMLTemplates\2016 Datasheet - 1 Column.xml");


            void datasheetCreator(string ID, string textToAppend)
            {
                datasheet.SelectSingleNode("//*[@id='"+ ID + "']").InnerText = textToAppend;
            }

            datasheetCreator("TODAYDATE", DateTime.Now.ToString("MM/dd/yyyy"));
            datasheetCreator("SOFTVERSION", _softversion);
            datasheetCreator("PROJECTNAME", _name);
            datasheetCreator("CITY", _city);
            datasheetCreator("CLIMATEZONE", _climateZone);
            datasheetCreator("PHOTO", (property.photovoltaic == "0") ? "N/A" : (property.photovoltaic.Length == 1) ? property.photovoltaic + ".00 kWdc" : property.photovoltaic + " kWdc");
            datasheetCreator("HERS", "N/A");
            datasheetCreator("WALLTYPE", _sidingOrStucco);
            datasheetCreator("PLANNAME", _name);
            datasheetCreator("FILENAME", property.fileName);
            datasheetCreator("SQFT", property.squareFeet);
            datasheetCreator("ABVP", _aboveCodePerc);
            datasheetCreator("COOLP", _spaceCool);
            datasheetCreator("STORIES", property.stories);
            datasheetCreator("GLAZINGP", (property.glazing == 0) ? "0.00%" : property.glazing.ToString().Split('.')?[1]?.Insert(2, ".") + "%");
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
            datasheetCreator("AFUE", property.afue);
            datasheetCreator("DUCTINS", "R-" + property.ductInsul);
            datasheetCreator("WHF", _wholeHouseFan);
            datasheetCreator("FANWAT", (property.fanWattage == "1") ? "Yes" : "-");
            datasheetCreator("AIRFLOW", (property.airflow == "1") ? "Yes ("+ property.airflowVal + ")" : "-");
            datasheetCreator("DUCTTEST", (property.ductTestingReq == "1") ? "Yes (" + property.ductTestingVal + "%)" : "-");
            datasheetCreator("CFM", (_indoorAirQual != "0") ? "Yes (" + _indoorAirQual + ")" : "-");
            datasheetCreator("REFCHARGE", (property.refCharg == "1") ? "Yes" : "-");
            datasheetCreator("SEERVERIF", (property.seerVerif == "1") ? "Yes" : "-");
            datasheetCreator("EERVERIF", (property.eerVerif == "1") ? "Yes" : "-");
            datasheetCreator("INFILTRATION", (_insulConsQual == "Standard") ? "-" : "Yes ("+ _airLeakage + ")");
            datasheetCreator("DUCTCOND", _ductInConditioned);
            datasheetCreator("LOWLEAK", (property.lowLeakageAir == "Has Low Leakage Air Handler") ? "Yes" : "-");
            datasheetCreator("BURRIEDDUCT", _buriedDuct);
            datasheetCreator("SURFAREA", _surfaceArea);
            datasheetCreator("INSULINSPECT", (_Qii == "1") ? "Yes" : "-");
            datasheetCreator("FUELTYPE", property.fuelType);
            datasheetCreator("UEF", _waterHeater);
            datasheetCreator("DISTRIBUTION", property.distribution);
           
            datasheetCreator("WIN1", (windowTypes.ElementAtOrDefault(0) != null) ? windowTypes[0] : "");
            datasheetCreator("WIN2", (windowTypes.ElementAtOrDefault(1) != null) ? windowTypes[1] : "");
            datasheetCreator("WIN3", (windowTypes.ElementAtOrDefault(2) != null) ? windowTypes[2] : "");
            datasheetCreator("WIN4", (windowTypes.ElementAtOrDefault(3) != null) ? windowTypes[3] : "");
  
           
            datasheetCreator("UVAL1", (windowsUV.ElementAtOrDefault(0) != null) ? windowsUV[0] : "");
            datasheetCreator("UVAL2", (windowsUV.ElementAtOrDefault(1) != null) ? windowsUV[1] : "");
            datasheetCreator("UVAL3", (windowsUV.ElementAtOrDefault(2) != null) ? windowsUV[2] : "");
            datasheetCreator("UVAL4", (windowsUV.ElementAtOrDefault(3) != null) ? windowsUV[3] : "");

            
            datasheetCreator("SHGC1", (windowsSHGC.ElementAtOrDefault(0) != null) ? windowsSHGC[0] : "");
            datasheetCreator("SHGC2", (windowsSHGC.ElementAtOrDefault(1) != null) ? windowsSHGC[1] : "");
            datasheetCreator("SHGC3", (windowsSHGC.ElementAtOrDefault(2) != null) ? windowsSHGC[2] : "");
            datasheetCreator("SHGC4", (windowsSHGC.ElementAtOrDefault(3) != null) ? windowsSHGC[3] : "");
    

            datasheet.Save(_name + " - Datasheet.doc");
        }
    }
}
