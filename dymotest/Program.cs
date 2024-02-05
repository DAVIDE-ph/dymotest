using DymoSDK.Implementations;
using DymoSDK.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dymotest
{
    public class Program
    {
        #region Variable definitions
        public static IEnumerable<DymoSDK.Interfaces.IPrinter> _printers;
        public static DymoSDK.Interfaces.IDymoLabel dymoSDKLabel;
        public static DymoSDK.Interfaces.ILabelObject _selectedLabelObject;
        public static DymoSDK.Interfaces.IPrinter _selectedPrinter;
        public static string Model = "C:\\Users\\dp\\Downloads\\d2.dymo";
        public static string _selectedRoll;


        public static List<PrinterStruct> Stampanti = new List<PrinterStruct>();
        public static List<LabelStruct> Campi = new List<LabelStruct>();
        #endregion

        static void Main(string[] args)
        {
            string id ="";
            //init DYMO
            DymoSDK.App.Init();
            dymoSDKLabel = DymoLabel.Instance;
            // load printers
            LoadPrinters();

            # region make a list of printers not in use.This is example how to check if there is printers and get them info
            foreach (var printer in Printers) 
            {
                var st = new PrinterStruct();
                st.DriverName= printer.DriverName;
                st.IsAutoCutSupported = printer.IsAutoCutSupported;
                st.IsConnected= printer.IsConnected;
                st.IsLocal= printer.IsLocal;
                st.IsTwinTurbo= printer.IsTwinTurbo;
                st.Name= printer.Name;
                st.PrinterType= printer.PrinterType;
                Stampanti.Add(st);
                Console.WriteLine(printer.Name);

            }
            #endregion

            //select dymo File
            dymoSDKLabel.LoadLabelFromFilePath(Model);
            //get dymo labels
            dymoSDKLabel.GetPreviewLabel();
            LabelObjects = dymoSDKLabel.GetLabelObjects().ToList();

            #region get labels with name this is not in use.This is example how to check labels
            foreach (var labs in _labelObjects)
            {
                
                var cmp = new LabelStruct();
                cmp.Name= labs.Name;
                cmp.Type = labs.Type;
                Campi.Add(cmp);
                Console.WriteLine(labs.Name);
                
            }
            #endregion

            #region Check if there is New records to print
            string SQLconn1 = "Server=13.0.0.95;Database=VMSystem;TrustServerCertificate=True;User Id=VMS;Password=DPH.OVH@13;Encrypt=False;";
             SqlConnection conn = new SqlConnection(SQLconn1);
            try
            {
                conn.Open();
                string s1 = "SELECT* FROM Visitors WHERE stampato=0";
                SqlCommand cmd = new SqlCommand(s1,conn);
                SqlDataReader rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    #region Set Label Values. in my case there is 1 label and 1 qr code 
                    int cont = 0;
                    foreach (var cmp in dymoSDKLabel.GetLabelObjects())
                    {
                        //in this case first label is input label with user name
                        if (cont == 0)
                        {
                            SelectedLabelObject = cmp;
                            dymoSDKLabel.UpdateLabelObject(SelectedLabelObject, "Nome:" + rd.GetValue(1).ToString());
                        }
                        //in this case second label is input label with Qr code
                        if (cont == 1)
                        {
                            SelectedLabelObject = cmp;
                            dymoSDKLabel.UpdateLabelObject(SelectedLabelObject, rd.GetValue(0).ToString());
                            id = rd.GetValue(0).ToString();
                        }
                        //in this case third label is surname
                        if(cont==2)
                        {
                            SelectedLabelObject = cmp;
                            dymoSDKLabel.UpdateLabelObject(SelectedLabelObject, "Cognome:"+ rd.GetValue(2).ToString());
                        }
                        //in this case third label is surname
                        if (cont == 3)
                        {
                            SelectedLabelObject = cmp;
                            dymoSDKLabel.UpdateLabelObject(SelectedLabelObject, "In visita:" + rd.GetValue(3).ToString());
                        }
                        //in this case third label is surname
                        if (cont == 4)
                        {
                            SelectedLabelObject = cmp;
                            DateTime dt = rd.GetDateTime(4);
                            dymoSDKLabel.UpdateLabelObject(SelectedLabelObject, "ingresso:" + dt.ToShortDateString()+" " +dt.ToShortTimeString());
                        }

                        cont++;
                    }
                    #endregion

                    #region Select Printer. in this case i will select first printer but you can change this function
                    cont = 0;
                    foreach (var selprt in Printers)
                    {
                        if (cont == 0)
                        {
                            SelectedPrinter = selprt;

                            cont++;
                        }
                    }
                    #endregion

                    //select role  Only for Twin Turbo 450 printer, the choises are Auto,Left,Right 
                    _selectedRoll = "Auto";
                    //print label
                    PrintLabelAction();
                    Console.WriteLine("Press any key to exit");
                    Console.ReadKey();
                }
                Update(id);
            }
            catch (Exception ex) 
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                conn.Close();
            }
            #endregion

          
        }

        public static DymoSDK.Interfaces.IPrinter SelectedPrinter
        {
            get { return _selectedPrinter; }
            set
            {
                _selectedPrinter = value;
                
            }
        }
        public static DymoSDK.Interfaces.ILabelObject SelectedLabelObject
        {
            get { return _selectedLabelObject; }
            set
            {
                _selectedLabelObject = value;
               
            }
        }
        private static void PrintLabelAction()
        {
            int copies = 1;
            if (_selectedPrinter != null)
            {
                //Send to print
                if (_selectedPrinter.Name.Contains("Twin Turbo"))
                {
                    int rollSel = _selectedRoll == "Auto" ? 0 : _selectedRoll == "Left" ? 1 : 2;
                    DymoPrinter.Instance.PrintLabel(dymoSDKLabel, _selectedPrinter.Name, copies, rollSelected: rollSel);
                }
                else
                    DymoPrinter.Instance.PrintLabel(dymoSDKLabel, _selectedPrinter.Name, copies);

                //If the label contains counter objects
                //Update counter object and preview to show the incresead value of the counter
                var counterObjs = dymoSDKLabel.GetLabelObjects().Where(w => w.Type == DymoSDK.Interfaces.TypeObject.CounterObject).ToList();
                if (counterObjs.Count > 0)
                {
                    foreach (var obj in counterObjs)
                        dymoSDKLabel.UpdateLabelObject(obj, copies.ToString());
                    
                }
            }
        }
        public static List<DymoSDK.Interfaces.ILabelObject> _labelObjects;
        public static List<DymoSDK.Interfaces.ILabelObject> LabelObjects
        {
            get
            {
                if (_labelObjects == null)
                    _labelObjects = new List<DymoSDK.Interfaces.ILabelObject>();
                return _labelObjects;
            }
            set
            {
                _labelObjects = value;
                
            }
        }
        public static IEnumerable<DymoSDK.Interfaces.IPrinter> Printers
        {
            get
            {
                if (_printers == null)
                    _printers = new List<DymoSDK.Interfaces.IPrinter>();
                return _printers;
            }
            set
            {
                _printers = value;

            }
        }
        private static async void LoadPrinters()
        {
            Printers = await DymoPrinter.Instance.GetPrinters();
        }
        public  class PrinterStruct
        {
            public string DriverName { get; set; }
            public bool IsAutoCutSupported { get; set; }
            public bool IsConnected { get; set; }
            public bool IsLocal { get; set; }
            public bool IsTwinTurbo {  get; set; }
            public string Name { get; set; }
            public string PrinterType { get; set; }
        }
        public class LabelStruct
        {
            public string Name { get; set; }
            public TypeObject Type { get; set; }    
        }

        private static void Update(string id)
        {
            string SQLconn1 = "Server=13.0.0.95;Database=VMSystem;TrustServerCertificate=True;User Id=VMS;Password=DPH.OVH@13;Encrypt=False;";
            SqlConnection conn = new SqlConnection(SQLconn1);

            try
            {
                conn.Open();
                string s1 = "UPDATE Visitors set stampato=1 WHERE id=@id";
                SqlCommand cmd = new SqlCommand(s1, conn);
                cmd.Parameters.AddWithValue("@id", id);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }
    }
}
