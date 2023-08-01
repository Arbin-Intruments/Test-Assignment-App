using System.Text;
using System.Text.RegularExpressions;

namespace InputModel
{
    public class EditObj {
           
        public string MachID { get; set; }
        public string ogMachID { get; set; }

        public string MachVolt { get; set; }
        public string ogMachVolt { get; set; }
        
        public string TesterID { get; set; }
        public string ogTestID { get; set; }
        
        public string StatID { get; set; }
        public string ogStatID { get; set; }
        
        public string kindOfEdit { get; set; }
        public bool edited { get; set; }

    }

public static class Constants
        {
            public const int NUM_COLS = 8; // # of columns / voltages in the machine excel file
            public const int PLUGS_PER_STATION = 2; // # of columns / voltages in the machine excel file
            public const int TRANSPARENT = 16777215; // ARGB value of transparent on excel files

        }

        public enum Permission 
        {
            Supervisor = 2,
            QC = 1,
            Tester = 0
        }

        public class TesterInfo {

            public TesterInfo (string pw = "000000", Permission perm = Permission.Tester, HashSet<string> machines = null) {
                password = pw;
                permission = perm;
                machineIDs = (machines == null) ? new HashSet<string>() : machines;
                // numMachines = machineIDs.Count;
            }

            public string? password { get; set; }
            public Permission permission { get; set; }
            public HashSet<string> machineIDs { get; set; }
            // public int numMachines { get; private set; }

        }

        public enum Voltage 
        {
            V110 = 110,  
            V120 = 120,
            V208 = 208,            
            V220 = 220,
            V380 = 380,
            V480 = 480,

        }

         public class StationInfo {

            public StationInfo (string id = "N99", bool source = false, int nPlugs = 0) {
                ID = id;
                isSource = source;
                numMachines = 0;
                machineIDs = new HashSet<string>();
                numPlugs = nPlugs;
                voltages = new SortedSet<Voltage>();
                plugsAndBoxes = new HashSet<List<string>>();
            }

            public StationInfo (string id, bool source, int nMachines, HashSet<string> machIDs, int nPlugs, SortedSet<Voltage> v, HashSet<List<string>> s) {
                ID = id;
                isSource = source;
                numMachines = nMachines;
                machineIDs = machIDs;
                numPlugs = nPlugs;
                voltages = v;
                plugsAndBoxes = s;

            }

            public string ID { get; set; }
            public bool isSource { get; set; }
            public int numMachines { get; set; }
            public HashSet<string> machineIDs { get; set; }
            public int numPlugs { get; set; }
            public SortedSet<Voltage>? voltages { get; set; }

            public HashSet<List<string>>? plugsAndBoxes { get; set; }
            // hashset of list of strings for each voltage used for front-end styling
            // each list in in the format: [ <path-to-image>, <alt-text>, <class-name> ]
            //                             [ "/bin/voltage-icons/380V-3P-Box.png", "380V-3P-Box", "v380V-3P-Box" ]

        }

         public enum Status 
        {
            Testing = 1,
            Waiting = 0,
        }

        public class MachineInfo {

            public MachineInfo () {
                voltage = 110;
                requiredStatVolt = Voltage.V110;
                // tester is nullable
                // station is nullable
                status = Status.Waiting;

            }

            public MachineInfo (int v, string tID, string sID, Status stat = Status.Testing) {
                voltage = (v != 0) ? v : 110;
                requiredStatVolt = getRequiredStatVolt(v);
                testerID = tID;
                stationID = sID;
                status = stat;

            }

            public int voltage { get; set; }
            public Voltage requiredStatVolt { get; set; }
            public string? testerID { get; set; }
            public string? stationID { get; set; }
            public Status status { get; set; }

            public void assign(string tID, string sID) {
                testerID = tID;
                stationID = sID;
                status = Status.Testing;
            }

            private Voltage getRequiredStatVolt (int machVolt) {

                int closestSmaller = 110;
                foreach (int v in Enum.GetValues(typeof(Voltage))) {
                    // if machine's voltage is one of the station voltages then closest match is that
                    if (machVolt == v) { 
                        return (Voltage)Enum.ToObject(typeof(Voltage), v); 
                    }
                    
                    // if machine's voltage is not one of the station voltages, then get the smallest one below
                    // works because voltage enum is looped smallest to largest
                    else if (v > machVolt) {
                        return (Voltage)Enum.ToObject(typeof(Voltage), closestSmaller);
                    }
                    else {
                        closestSmaller = v;
                    }
                }

                return Voltage.V110; // the machine voltage is lower for some reason
            }

        }


}