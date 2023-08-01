using System.Text;

namespace InputModel
{

    public abstract class InputModel
    {

        // TODO: connect with database
        public Dictionary<string, MachineInfo>? machineList { get; set; } // Machine_TestStation.txt

        public Dictionary<string, Permission>? taskTypes { get; set; } // TaskID_Permission.txt
        public Dictionary<string, TesterInfo>? testerList { get; set; } // Tester.txt
        public Dictionary<char, List<StationInfo>>? stationList { get; set; } // Machine Tracking Tables.xlsx
    
        // debug
        private const string ConfigPath = "Config.txt";
        public static string EmbeddedFolderPath => $@"{File.ReadAllText(ConfigPath)}";

        // helper function
        public static Voltage stringToVolt (string v) {
            switch (v)
            {
            case "110V":
                return Voltage.V110;
            case "120V":
                return Voltage.V120;
            case "208V":
                return Voltage.V208;
            case "220V":
                return Voltage.V220;
            case "380V":
                return Voltage.V380;
            case "480V":
                return Voltage.V480;
            default:
                return Voltage.V110;

            }
        }

        // populate member variables
        public virtual void getTaskTypes() {
            foreach (string line in File.ReadLines(Path.Combine(EmbeddedFolderPath, "Input", "TaskID_Permission.txt"), Encoding.UTF8))
            {
                // process the line
                // 1S,Tester

                string[] tParams = line.Split(',');
                
                string taskID = tParams[0];
                string perm_NotEnum = tParams[1];

                Permission perm;
                Enum.TryParse<Permission>(perm_NotEnum, out perm);

                if (taskTypes == null) {
                    taskTypes = new Dictionary<string, Permission>();
                    taskTypes.Add(taskID, perm);
                }
                else if(taskTypes.ContainsKey(taskID)){
                    taskTypes[taskID] = perm;
                }
                else {
                    taskTypes.Add(taskID, perm);
                }
            }
        }

        // get the member variable data
        public virtual void getTesterList() {}
        public virtual void getStationList() {}
        public virtual void getMachineList() {}


        // update member variables
        public virtual void updateStation(string newStationID, string oldStationID, string machineID) {
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

                setMachineList();
                setStationList();
               
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
        public virtual void updateTester(string newTesterID, string oldTesterID, string machineID) {
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

                setTesterList();
                setMachineList();
            
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


        // set member variables' respective file/db
        public virtual void setMachineList() {}
        public virtual void setTesterList() {}
        public virtual void setStationList() {}

      
        
    }
}

