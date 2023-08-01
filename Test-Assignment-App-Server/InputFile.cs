using ClosedXML.Excel;
using System.Text;
using System.Text.RegularExpressions;

namespace InputModel
{

    public class FileInputModel : InputModel
    {
        public FileInputModel() {

            Console.WriteLine("entered contstructor using file input");

            // get = update our lists in C#
            // set = update our lists in the file/DB

            getTaskTypes(); 

            getTesterList();

            getStationList();

            getMachineList();

        }


        // POPULATE MEMBER VARIABLES FROM FILES
        // uses abstract's getTaskList()

        public override void getTesterList() {
            string fileName = Path.Combine(EmbeddedFolderPath, "Input", "Input.xlsx");
            var wb = new XLWorkbook(fileName);
            IXLWorksheet ws = wb.Worksheet("tester_input"); // Put your sheet name here

            int r = 2; // skips header row
            while (!ws.Row(r).IsEmpty())
            {
                // process the line
                // tester ID	 password	permission	machine IDs
                // Ryan J.	            0	permission	
                // Joe	                0	permission	
                // Ryan M.	            0	permission	
                // Xia	                0	permission	224159-A
                // Gie	                0	permission	220767-A, 220767-B                                    
                
                if (testerList == null) {
                    testerList = new Dictionary<string, TesterInfo>();
                }

                var cell = ws.Cell(r, 1); // tester col 
                string testerID = cell.GetValue<string>();

                cell = ws.Cell(r, 2); // password col 
                string password = cell.GetValue<string>();

                cell = ws.Cell(r, 3); // permission col 
                string perm_NotEnum = cell.GetValue<string>();
                Permission perm;
                Enum.TryParse<Permission>(perm_NotEnum, out perm);

                cell = ws.Cell(r, 4); // machine ID col 
                string machineIDs_raw = cell.GetValue<string>();
                HashSet<string> machineIDs = new HashSet<string>();

                if (!cell.IsEmpty()) {
                    string[] machineIDs_strings = machineIDs_raw.Split(',');

                    // filters for duplicates and removes extra characters
                    foreach (string machineID_string in machineIDs_strings) {
                        if (machineID_string != "") {
                            machineIDs.Add(Regex.Replace(machineID_string, @"[^0-9a-zA-Z\-]+", ""));
                        }
                    }
                }


                TesterInfo testerInfo = new TesterInfo(password, perm, machineIDs);


                if(testerList.ContainsKey(testerID)) { // tester has already been added to the list
                    continue;
                }
                else {
                    testerList.Add(testerID, testerInfo);
                }

                ++r;

            }
        }

        public override void getStationList() {
            string fileName = Path.Combine(EmbeddedFolderPath, "Input", "Input.xlsx");
            var wb = new XLWorkbook(fileName);
            IXLWorksheet ws = wb.Worksheet("station_input"); // Put your sheet name here

            int r = 2; // skips header row
            while (!ws.Row(r).IsEmpty())
            {
                // process the line
                //     	480V-3P-Box	| 380V-3P-Box | 208V-3P-Box | 208V-3P-Plug  | 220V-1P-Plug  | 208V-1P-Plug	| 120V-1P-Plug | number of machines
                // A01		                380V    					120V                                                            2
                // A02		                380V	    				120V                                                            1
                // A03		                380V		    			120V                                                
                // A04		                380V			    		120V
                // A05		                380V				    	120V
                // A06			                            208V				
                // B01				                        208V	                     220V		                                    1
                
                if (stationList == null) {
                    stationList = new Dictionary<char, List<StationInfo>>();
                }

                var cell = ws.Cell(r, 1); // HEADER COL   
                string stationID = cell.GetValue<string>();

                char row = stationID[0]; // first character is the row station
                
                bool isSource = true;
                SortedSet<Voltage> voltages = new SortedSet<Voltage>();
                HashSet<List<string>> plugsAndBoxes = new HashSet<List<string>>();


                // c = 2 to skip header column
                for (int c = 2; c < Constants.NUM_COLS+1; ++c) {
                    cell = ws.Cell(r, c);
                    var cell_header = ws.Cell(1, c); // header row

                    if (cell.IsEmpty()) {
                        continue;
                    }
                    // Console.Write(cell.GetText() + ", ");

                    int backgroundColor = ws.Cell(r, c).Style.Fill.BackgroundColor.Color.ToArgb(); // Gives the Argb values for the cell's background. Replace rowIndex and colIndex with the index numbers of the cell you want to check
                    isSource = backgroundColor != Constants.TRANSPARENT;

                    string volt_NotEnum = cell.GetValue<string>();
                    Voltage voltage = stringToVolt(volt_NotEnum);

                    voltages.Add(voltage);


                    string plugOrBox = cell_header.GetValue<string>();
                    List<string> altAndSrc = new List<string>();
                    altAndSrc.Add(Path.Combine("Assets", "voltage_icons", plugOrBox + ".png"));
                    // altAndSrc.Add(@$"{Directory.GetParent(Directory.GetCurrentDirectory())}");
                    
                    altAndSrc.Add(plugOrBox);
                    altAndSrc.Add("v" + plugOrBox);

                    plugsAndBoxes.Add(altAndSrc);
                }

                cell = ws.Cell(r, Constants.NUM_COLS + 2); // machine ID col 
                string machineIDs_raw = cell.GetValue<string>();
                HashSet<string> machineIDs = new HashSet<string>();

                if (!cell.IsEmpty()) {
                    string[] machineIDs_strings = machineIDs_raw.Split(',');

                    // filters for duplicates and removes extra characters
                    foreach (string machineID_string in machineIDs_strings) {
                        string machineID = Regex.Replace(machineID_string, @"[^0-9a-zA-Z\-]+", "");
                        if (machineID != "" && !machineIDs.Contains(machineID)) {
                            machineIDs.Add(machineID);
                        }
                    }
                }

                int numMachines = machineIDs.Count;

                StationInfo stationInfo = new StationInfo(stationID, isSource, numMachines, machineIDs, Constants.PLUGS_PER_STATION, voltages, plugsAndBoxes);


                if(stationList.ContainsKey(row)) {
                    stationList[row].Add(stationInfo);
                }
                else {
                    List<StationInfo> rowsStations = new List<StationInfo>();
                    rowsStations.Add(stationInfo);
                    stationList.Add(row, rowsStations);
                }

                ++r;

            }
        }

        public override void getMachineList() {
            string fileName = Path.Combine(EmbeddedFolderPath, "Input", "Input.xlsx");
            var wb = new XLWorkbook(fileName);
            IXLWorksheet ws = wb.Worksheet("machine_input"); // Put your sheet name here

            int r = 2; // skips header row
            while (!ws.Row(r).IsEmpty())
            {
                // process the line
                // machine ID	|   machine voltage |   tester ID	|   station ID
                // 220759-A	                480V		                    Z01
                // 220759-B	                480V	    Xiaomei	            A02
                // 220767-A	                480V	    Gie	                D02
                // 220767-B	                480V	    Gie	                C01
                // 224087-A	                480V	    Song            	B01
                // 224135-A	                480V		                    A02
                // 224135-B	                480V		
                // 224159-A	                480V	    Xia	                                  
                                
                if (machineList == null) {
                    machineList = new Dictionary<string, MachineInfo>();
                }

                var cell = ws.Cell(r, 1); // machine col 
                string machineID = cell.GetValue<string>();

                cell = ws.Cell(r, 2); // machine voltage col 
                string voltage_string = cell.GetValue<string>();

                string voltageSymbols = "[vV]";
                voltage_string = Regex.Replace(voltage_string, voltageSymbols, String.Empty);
                int voltage = Convert.ToInt32(voltage_string);

                cell = ws.Cell(r, 3); // tester ID col 
                string testerID = cell.GetValue<string>();

                cell = ws.Cell(r, 4); // station ID col 
                string stationID = cell.GetValue<string>();

                MachineInfo machineInfo = new MachineInfo(voltage, testerID, stationID);

                // Console.WriteLine($"{machineID}: <{voltage}, {testerID}, {stationID}>");

                if(machineList.ContainsKey(machineID)) { // tester has already been added to the list
                    continue;
                }
                else {
                    machineList.Add(machineID, machineInfo);
                }

                ++r;

            }
        }

       
        // UPDATE MEMBER VARIABLES
        // taken from abstract class


        // SET FILES
        // edit "machine_list"
        public override void setMachineList() {
            
            // process the line
            // machine ID	|   machine voltage |   tester ID	|   station ID
            // 220759-A	                480V		                    Z01
            // 220759-B	                480V	    Xiaomei	            A02
            // 220767-A	                480V	    Gie	                D02
            // 220767-B	                480V	    Gie	                C01
            // 224087-A	                480V	    Song            	B01
            // 224135-A	                480V		                    A02
            // 224135-B	                480V		
            // 224159-A	                480V	    Xia	                               

            // unedited sheets
            string fileName = Path.Combine(EmbeddedFolderPath, "Input", "Input.xlsx");
            var wb = new XLWorkbook(fileName);
            IXLWorksheet ws_tester = wb.Worksheet("tester_input"); // tester input
            IXLWorksheet ws_station = wb.Worksheet("station_input"); // station input

            var workbook = new XLWorkbook();
            
            IXLWorksheet ws_machine = workbook.Worksheets.Add("machine_input"); // machine input
            
            // add header
            // machine ID	machine voltage	tester ID	station ID
            ws_machine.Cell("A1").Value = "machine ID";
            ws_machine.Cell("B1").Value = "machine voltage";
            ws_machine.Cell("C1").Value = "tester ID";
            ws_machine.Cell("D1").Value = "station ID";

            if (machineList != null) {
                int r = 2;
                foreach (var machine in machineList) {
                    // machine ID	machine voltage	tester ID	station ID
                    ws_machine.Cell($"A{r}").Value = machine.Key;
                    ws_machine.Cell($"B{r}").Value = machine.Value.voltage;
                    ws_machine.Cell($"C{r}").Value = machine.Value.testerID;
                    ws_machine.Cell($"D{r}").Value = machine.Value.stationID;
                    
                    ++r;
                }
            }    

            // add pre-exisitng pages
            workbook.AddWorksheet(ws_tester);     
            workbook.AddWorksheet(ws_station);

            string newFileName = Path.Combine(EmbeddedFolderPath, "Input", "Input.xlsx");
            workbook.SaveAs(newFileName);
        }

        // edit "tester_input"
        public override void setTesterList() {
            
            // process the line
            // tester ID	 password	permission	machine IDs
            // Ryan J.	            0	permission	
            // Joe	                0	permission	
            // Ryan M.	            0	permission	
            // Xia	                0	permission	224159-A
            // Gie	                0	permission	220767-A, 220767-B     

            // unedited sheets
            string fileName = Path.Combine(EmbeddedFolderPath, "Input", "Input.xlsx");
            var wb = new XLWorkbook(fileName);
            IXLWorksheet ws_machine = wb.Worksheet("machine_input"); // machine input
            IXLWorksheet ws_station = wb.Worksheet("station_input"); // station input

            var workbook = new XLWorkbook();

            // add pre-exisitng pages
            workbook.AddWorksheet(ws_machine);   

            IXLWorksheet ws_tester = workbook.Worksheets.Add("tester_input"); // tester input
            
            // add header
            // tester ID	 password	permission	machine IDs
            ws_tester.Cell("A1").Value = "tester ID";
            ws_tester.Cell("B1").Value = "password";
            ws_tester.Cell("C1").Value = "permissions";
            ws_tester.Cell("D1").Value = "machine IDs";

            if (testerList != null) {
                int r = 2;
                foreach (var tester in testerList) {
                    Console.WriteLine($">> Tester: {tester.Key} - {tester.Value.password} {Enum.GetName(tester.Value.permission)}");

                    // tester ID	 password	permission	machine IDs
                    ws_tester.Cell($"A{r}").Value = tester.Key;
                    ws_tester.Cell($"B{r}").Value = tester.Value.password;
                    ws_tester.Cell($"C{r}").Value = Enum.GetName(tester.Value.permission);

                    string machineIDs_string = "";
                    if (tester.Value.machineIDs != null && tester.Value.machineIDs.Count > 0) {
                        foreach (string machineID in tester.Value.machineIDs) {
                            if (machineID != "") {
                                machineIDs_string += machineID + ",";
                            }
                        }
                    }
                    
                    ws_tester.Cell($"D{r}").Value = machineIDs_string;
                    
                    ++r;
                }
            }    

            // add pre-exisitng pages
            workbook.AddWorksheet(ws_station);

            string newFileName = Path.Combine(EmbeddedFolderPath, "Input", "Input.xlsx");
            workbook.SaveAs(newFileName);
        }

        // set "station_input"
         public override void setStationList() {
            
            // process the line
                //     	480V-3P-Box	| 380V-3P-Box | 208V-3P-Box | 208V-3P-Plug  | 220V-1P-Plug  | 208V-1P-Plug	| 120V-1P-Plug | number of machines
                // A01		                380V    					120V                                                            2
                // A02		                380V	    				120V                                                            1
                // A03		                380V		    			120V                                                
                // A04		                380V			    		120V
                // A05		                380V				    	120V
                // A06			                            208V				
                // B01				                        208V	                     220V		

            // unedited sheets
            string fileName = Path.Combine(EmbeddedFolderPath, "Input", "Input.xlsx");
            var wb = new XLWorkbook(fileName);
            IXLWorksheet ws_machine = wb.Worksheet("machine_input"); // machine input
            IXLWorksheet ws_tester = wb.Worksheet("tester_input"); // tester input
            IXLWorksheet ws_station = wb.Worksheet("station_input"); // station input

            var workbook = new XLWorkbook();

            if (stationList != null) {
                int r = 2;
                foreach (var row in stationList) {
                    foreach (var station in row.Value) {
                        // 	480V-3P-Box	380V-3P-Box	208V-3P-Box	208V-3P-Plug	220V-1P-Plug	208V-1P-Plug	120V-1P-Plug	Machine IDs
                        ws_station.Cell($"I{r}").Value = station.numMachines;

                        string machineIDs_string = "";
                        foreach (string machineID in station.machineIDs) {
                            if (machineID != "") {
                                machineIDs_string += machineID + ",";
                            }
                        }
                        ws_station.Cell($"J{r}").Value = machineIDs_string;
                        ++r;
                    }
                }
            }    

            // add pre-exisitng pages
            workbook.AddWorksheet(ws_machine);     
            workbook.AddWorksheet(ws_tester);     
            workbook.AddWorksheet(ws_station);

            string newFileName = Path.Combine(EmbeddedFolderPath, "Input", "Input.xlsx");
            workbook.SaveAs(newFileName);
        }
       
    }
}