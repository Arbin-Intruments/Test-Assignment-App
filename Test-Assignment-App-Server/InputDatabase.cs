using System;  
using ClosedXML.Excel;
using System.Data.OleDb;  
using System.Globalization;
using System.Text.RegularExpressions;

namespace InputModel
{  
    public class DBInputModel : InputModel {  

        string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + EmbeddedFolderPath + @"\Input\ArbinDataCenter.mdb";
        // string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\File1\Software-Testing$\Production apps\Test Station Assignment Recorder v1.0\Input\Input\ArbinDataCenter.mdb";

        public DBInputModel() {
            Console.WriteLine("entered contstructor using DB input");
            Console.WriteLine(connectionString);

            Fix(); 

            // get = update our lists in C#
            // set = update our lists in the file/DB

            getTaskTypes(); 

            // will not have all the information
            getTesterList();

            // will not have all the information
            getStationList();

            // update the missing information
            getMachineList();

        }

        public DBInputModel(string connString) {
            connectionString = connString;

            Console.WriteLine("entered contstructor using DB input");

            // get = update our lists in C#
            // set = update our lists in the file/DB

            getTaskTypes(); 

            getTesterList();

            getStationList();

            getMachineList();
        }

        public override void getTesterList() {  

            // public string? password { get; set; }
            // public Permission permission { get; set; }
            // public HashSet<string> machineIDs { get; set; }

            if (testerList == null) {
                testerList = new Dictionary<string, TesterInfo>();
            }

            using(OleDbConnection connection = new OleDbConnection(connectionString)) {  
                try {  
                    string strSQL = "SELECT * FROM EmployeeList WHERE Title='Testing Technician' OR Title='Testing' OR Title='Tester' OR Title='Testing Engineer'";  
                    OleDbCommand command = new OleDbCommand(strSQL, connection);  

                    connection.Open();  

                    using(OleDbDataReader reader = command.ExecuteReader()) {  
                        // Console.WriteLine("------------Original data----------------");  
                        while (reader.Read()) {  

                            string? testerID = reader["Name"].ToString().Replace(' ', '-');
                            testerID += "-" + reader["InternalID"].ToString();
                            string? password = reader["InternalID"].ToString() + "";
                            Permission permission = Permission.Tester;
                            HashSet<string> machineIDs = new HashSet<string>();

                            // Console.WriteLine("{0} - {1}", testerID, password);  
                        
                            TesterInfo newTester = new TesterInfo(password, permission, machineIDs);
                            testerList.Add(testerID, newTester);
                        }
                        reader.Close();  
                    }  

                } catch (Exception ex) {  
                    Console.WriteLine(ex.Message);  
                }  
            }  
        }  

        public override void getMachineList() {  

            // public int voltage { get; set; }
            // public string? testerID { get; set; }
            // public string? stationID { get; set; }
            // public Status status { get; set; }

            if (machineList == null) {
                machineList = new Dictionary<string, MachineInfo>();
            }

            DateTime localDate = DateTime.Now.Date.AddDays(-1);
            var culture = new CultureInfo("en-US");

            using(OleDbConnection connection = new OleDbConnection(connectionString)) {  
                try {  
                    string strSQL = "SELECT * FROM CopyOfSNTbl " +
                                    "WHERE Stage='' OR Stage='QC - USA' OR Stage='Planning' OR " +
                                    "Stage='Testing/QC - USA' OR Stage='Assembly - USA' OR " +
                                    "Stage='Assembly - China' OR Stage='QC - China' " + 
                                    "ORDER BY NewETA desc ";  
                    OleDbCommand command = new OleDbCommand(strSQL, connection);  

                    connection.Open();  

                    using(OleDbDataReader reader = command.ExecuteReader()) {  
                        // Console.WriteLine("------------{0} Original data----------------", localDate.ToString(culture));  
                        while (reader.Read()) {  

                            string? machineID = reader["SN"].ToString();

                            string voltage_string = (reader["AC_Power"].ToString() == "") ? "110" : reader["AC_Power"].ToString();

                            // string text = "415V-3P";
                            string pattern = @"\d*";
                            Match match = Regex.Match(voltage_string, pattern);
                            if (match.Success) {
                                // Console.WriteLine("{0} at {1}", match.Value, match.Index);
                                voltage_string = match.Value;
                            }

                            int voltage = Int32.Parse(voltage_string);

                            string? date_string = reader["NewETA"].ToString() == "" ? localDate.ToString(culture) : reader["NewETA"].ToString();
                            DateTime date = DateTime.Parse(date_string, culture);

                            string? testerID = reader["Tester"].ToString().TrimStart();
                            string? stationID = reader["Test_Station"].ToString();

                            Status status = (Regex.IsMatch(reader["Stage"].ToString(), @"^.*QC.*$")) ?
                                Status.Testing : Status.Waiting;

                            // filter if machineID doesn't match format
                            if ( !Regex.IsMatch(machineID, @"[\d]{6}[-]{0,1}[\w]{0,1}$") ) {
                                continue;
                            }
                            
                            // Permission permission = Permission.Tester;
                            // HashSet<string> machineIDs = new HashSet<string>();
                            if (date <= localDate) {
                                Console.WriteLine("-- {0}  vs  {1}", date_string, localDate.ToString(culture));
                                break;
                            }

                            // Console.WriteLine("{0} - {1}V {2}", machineID, voltage, date);  
                            // Console.WriteLine("{0} - [{1}V] {2} @ {3} - {4}", machineID, voltage, testerID, stationID, date);  
                        
                            MachineInfo newMachine = new MachineInfo(voltage, testerID, stationID, status);
                            machineList.Add(machineID, newMachine);


                            // GET MISSING INFO FOR TESTER AND MACHINES
                            char stationRow = 'N';
                            int stationPos = -1;

                            if (stationID.Length >= 2) {
                                stationRow = char.Parse(stationID.Substring(0, 1)); // ex. A
                                stationPos = Int32.Parse(stationID.Substring(1)); // ex. 3
                            }

                            if (stationID != "" && (stationList[stationRow].Count > stationPos && stationPos <= 0)) {
                                updateStation(stationID, "None", machineID);
                            }

                            if (testerList.ContainsKey(testerID) && testerID != "") {
                                updateTester(testerID, "None", machineID);
                            }


                        }
                        reader.Close();  
                    }  

                } catch (Exception ex) {  
                    Console.WriteLine(ex.Message);  
                }  
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

                    int backgroundColor = ws.Cell(r, c).Style.Fill.BackgroundColor.Color.ToArgb(); // Gives the Argb values for the cell's background. Replace rowIndex and colIndex with the index numbers of the cell you want to check
                    isSource = backgroundColor != Constants.TRANSPARENT;

                    string volt_NotEnum = cell.GetValue<string>();
                    Voltage voltage = stringToVolt(volt_NotEnum);

                    voltages.Add(voltage);


                    string plugOrBox = cell_header.GetValue<string>();
                    List<string> altAndSrc = new List<string>();
                    altAndSrc.Add(Path.Combine("Assets", "voltage_icons", plugOrBox + ".png"));
                    
                    altAndSrc.Add(plugOrBox);
                    altAndSrc.Add("v" + plugOrBox);

                    plugsAndBoxes.Add(altAndSrc);
                }

                cell = ws.Cell(r, Constants.NUM_COLS + 2); // machine ID col 
                string machineIDs_raw = cell.GetValue<string>();
                HashSet<string> machineIDs = new HashSet<string>();

                // does not get the machineIDs or numMachines
                // will be done in getMachineList

                StationInfo stationInfo = new StationInfo(stationID, isSource, 0, machineIDs, Constants.PLUGS_PER_STATION, voltages, plugsAndBoxes);


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


        // UPDATE TESTER AND STATION
        public override void updateStation(string newStationID, string oldStationID, string machineID) {

            Console.WriteLine("{0} - replacing {1} with {2}", machineID, oldStationID, newStationID);

            // newStationID would be ex. A03
            char newStationRow = char.Parse(newStationID.Substring(0, 1)); // ex. A
            int newStationPos = Int32.Parse(newStationID.Substring(1)); // ex. 3

            // oldStationID would be ex. A03
            char oldStationRow = 'N';
            int oldStationPos = -1;
            if (oldStationID != "None" && oldStationID != "") {
                oldStationRow = char.Parse(oldStationID.Substring(0, 1)); // ex. A
                oldStationPos = Int32.Parse(oldStationID.Substring(1)); // ex. 3
            }
            

            if (machineList != null && stationList != null) {
                machineList[machineID].stationID = newStationID;

                Console.WriteLine($"Station {newStationID}: {stationList[newStationRow].ElementAt(newStationPos-1).numMachines}");
                if (oldStationID != "None" && oldStationID != "") {
                    Console.WriteLine($"Old Station {oldStationID}: {stationList[oldStationRow].ElementAt(oldStationPos-1).numMachines}");
                }
                else {
                    Console.WriteLine("Old Station None: 0");
                }
                Console.WriteLine("----");

                machineList[machineID].stationID = newStationID;
                Console.WriteLine("CHECKING: " + stationList[newStationRow].ElementAt(newStationPos-1).ID);
                // ++(stationList[newStationRow].ElementAt(newStationPos-1).numMachines);
                if (!stationList[newStationRow].ElementAt(newStationPos-1).machineIDs.Contains(machineID)) {
                    stationList[newStationRow].ElementAt(newStationPos-1).machineIDs.Add(machineID);
                }
                stationList[newStationRow].ElementAt(newStationPos-1).numMachines = stationList[newStationRow].ElementAt(newStationPos-1).machineIDs.Count;

                if (oldStationID != "" && oldStationID != "None") {
                    // --(stationList[oldStationRow].ElementAt(oldStationPos-1).numMachines);
                    stationList[oldStationRow].ElementAt(oldStationPos-1).machineIDs.Remove(machineID);
                    stationList[oldStationRow].ElementAt(oldStationPos-1).numMachines = stationList[newStationRow].ElementAt(newStationPos-1).machineIDs.Count;

                }

                setMachineList(machineID, machineList[machineID].testerID, newStationID);

                Console.WriteLine($"Station {newStationID}: {stationList[newStationRow].ElementAt(newStationPos-1).numMachines}");
                if (oldStationID != "None" && oldStationID != "") {
                    Console.WriteLine($"Old Station {oldStationID}: {stationList[oldStationRow].ElementAt(oldStationPos-1).numMachines}");
                }
                else {
                    Console.WriteLine("Old Station None: 0");
                }
            
            }
            else {
                if (machineList != null ) Console.WriteLine("ERROR: Cannot find machine " + machineID);
                if (stationList != null ) Console.WriteLine("ERROR: Cannot find station " + newStationID);
            
            }
        }
        public override void updateTester(string newTesterID, string oldTesterID, string machineID) {

            Console.WriteLine("{0} - replacing {1} with {2}", machineID, oldTesterID, newTesterID);

            if (machineList != null && testerList != null) {
                machineList[machineID].testerID = newTesterID;

                Console.Write($"\nTester {newTesterID}: ");
                foreach (string m in testerList[newTesterID].machineIDs) {
                    Console.Write($"{m} ");
                }
                Console.WriteLine();
                if (oldTesterID != "None" && oldTesterID != "") {
                    Console.Write($"Old Tester {oldTesterID}: ");
                    foreach (string m in testerList[oldTesterID].machineIDs) {
                        Console.Write($"{m} ");
                    }
                    Console.WriteLine();
                }
                else {
                    Console.WriteLine("Old Tester None: ");
                }

                /////
                Console.WriteLine($"\nAdded {machineID} to {newTesterID}\n");
                testerList[newTesterID].machineIDs.Add(machineID);
                
                if (oldTesterID != "" && oldTesterID != "None") {
                    testerList[oldTesterID].machineIDs.Remove(machineID);
                }
                /////

                setMachineList(machineID, newTesterID, machineList[machineID].stationID);
            
                Console.Write($"Tester {newTesterID}: ");
                foreach (string m in testerList[newTesterID].machineIDs) {
                    Console.Write($"{m} ");
                }
                if (oldTesterID != "None" && oldTesterID != "") {
                    Console.Write($"Old Tester {oldTesterID}: ");
                    foreach (string m in testerList[oldTesterID].machineIDs) {
                        Console.Write($"{m} ");
                    }
                    Console.WriteLine();
                }
                else {
                    Console.WriteLine("Old Tester None: ");
                }
            }
            else {
                if (machineList != null ) Console.WriteLine("ERROR: Cannot find machine " + machineID);
                if (testerList != null ) Console.WriteLine("ERROR: Cannot find tester " + newTesterID);
            
            }
        }




        public void setMachineList(string machineID, string newTester, string newStation) {

            using(OleDbConnection connection = new OleDbConnection(connectionString)) {  
                try {  
                    connection.Open();

                    string strSQL = "UPDATE CopyOfSNTbl SET Tester='" + newTester + "' WHERE SN='" + machineID + "' ";

                    OleDbCommand command = new OleDbCommand(strSQL, connection);  
                    command.ExecuteReader();  

                    strSQL = "UPDATE CopyOfSNTbl SET Test_Station='" + newStation + "' WHERE SN='" + machineID + "' ";
                    
                    command = new OleDbCommand(strSQL, connection);  
                    command.ExecuteReader();  

                    connection.Close();

                } catch (Exception ex) {  
                    Console.WriteLine(ex.Message);  
                }  
            
            }
        }


        private void Fix () {

            using(OleDbConnection connection = new OleDbConnection(connectionString)) {  
                try {  
                    connection.Open();

                    string strSQL = "UPDATE CopyOfSNTbl SET Tester='' WHERE Tester='Feng-lin#49' ";

                    OleDbCommand command = new OleDbCommand(strSQL, connection);  
                    command.ExecuteReader();  

                    // strSQL = "UPDATE CopyOfSNTbl SET Tester='' WHERE Tester='Evan Johnson#79593' ";
                    
                    // command = new OleDbCommand(strSQL, connection);  
                    // command.ExecuteReader();  

                    connection.Close();

                } catch (Exception ex) {  
                    Console.WriteLine(ex.Message);  
                }  
            
            }

        }

          // // Add a new row  
                    // strSQL = "INSERT INTO Developer(Name, Address) VALUES ('New Developer', 'New Address')";
                    
                    // command = new OleDbCommand(strSQL, connection);  
                    // // Execute command  
                    // command.ExecuteReader();  

                    // // The following code snippet updates rows that match with the WHERE condition in the query.
                    // // Update rows  
                    // strSQL = "UPDATE Developer SET Name = 'Updated Name' WHERE Name = 'New Developer'";
                    // command = new OleDbCommand(strSQL, connection);  
                    // command.ExecuteReader();  

                    // // The following code snippet deletes rows that match with the WHERE condition.
                    // // Delete rows  
                    // strSQL = "DELETE FROM Developer WHERE Name = 'Updated Name'";  
                    // command = new OleDbCommand(strSQL, connection);  
                    // command.ExecuteReader();  

 
    }  
} 

