using System;
using System.Collections.Generic;
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
using OfficeOpenXml;
using Invoice;
using InvoiceDataTemplates;
using System.IO;

namespace GenerateInvoice
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

        private void btnPreviousTab_Click(object sender, RoutedEventArgs e)
        {
            var invoiceTemplate = InvoiceTemplate.DefaultTemplate();
            invoiceTemplate.ModifyDocument();
        }

        private void btnPreviousTab_Click2(object sender, RoutedEventArgs e)
        {
            var invoiceTemplate = InvoiceTemplate.DefaultTemplate();
            invoiceTemplate.ModifyDocument();
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            // ... A List.
            List<string> data = new List<string>();
            data.Add("WALK IN");
            data.Add("CONCORD");
            data.Add("CAPER");
            data.Add("Minar");

            // ... Get the ComboBox reference.
            var comboBox = sender as ComboBox;

            // ... Assign the ItemsSource to the List.
            comboBox.ItemsSource = data;

            // ... Make the first item selected.
            comboBox.SelectedIndex = 0;
        }
        private InvoiceTemplate createTemplateForSingleInvoice()
        {
            Dictionary<string, string> addresses = new Dictionary<string, string>
            {
                ["WALK IN"] = " ",
                ["CONCORD"] = "CONCORD EXOTIC VOYAGES(I) PVT.LTD. 407 / 408 - 4TH FLOOR, GERA IMPERIUM II, PATTO, PANJIM, GOA - 403001.",
                ["CAPER"] = "CAPER TRAVEL COMPANY PVT. LTD.HOUSE NO-14/242/B	BEACH PLAZA ANNEXE BUILDING, NEAR KAMAT KINARA, NOMOXIM,CARANZALEM,MIRAMAR, TISWADI, GOA-403002 ",
                ["Minar"] = "Minar Travels (India) Pvt Ltd 101/102, 1st Floor Gera's Imperium II, Patto Plaza Panjim Goa 403001"
            };

            Dictionary<string, string> gst = new Dictionary<string, string>
            {
                ["WALK IN"] = " ",
                ["CONCORD"] = "30AACCC1364LIZZ",
                ["CAPER"] = "30AABCC5600J1Z9",
                ["Minar"] = "30AAACM1267A1ZC"
            };

            string gt = gst[CompanyName.SelectedItem.ToString()];
            string adr = addresses[CompanyName.SelectedItem.ToString()];

            if(gt == " ")
            {
                gt = CompanyGST.Text != string.Empty ? CompanyGST.Text : " ";
                adr = CompanyAddress.Text != string.Empty ? CompanyAddress.Text : " ";
            }
            var company = new Company(CompanyName.SelectedItem.ToString(), gt, adr);
            List<Room> rooms = new List<Room>();
            List<Service> services = new List<Service>();
            if (Room1Detail.Text != string.Empty)
            {
                rooms.Add(new Room(Room1Detail.Text, double.Parse(Room1Rate.Text)));
            }
            if (Room2Detail.Text != string.Empty)
            {
                rooms.Add(new Room(Room2Detail.Text, double.Parse(Room2Rate.Text)));
            }
            if (Room3Detail.Text != string.Empty)
            {
                rooms.Add(new Room(Room3Detail.Text, double.Parse(Room3Rate.Text)));
            }
            if (Service1Detail.Text != string.Empty)
            {
                services.Add(new Service(Service1Detail.Text, double.Parse(Service1Rate.Text)));
            }
            if (Service2Detail.Text != string.Empty)
            {
                services.Add(new Service(Service2Detail.Text, double.Parse(Service2Rate.Text)));
            }

            var invoiceDetail = new InvoiceDetail(
                InvoiceDate.Text,
                InvoiceNo.Text,
                CheckInDate.Text,
                CheckOutDate.Text,
                NoOfGuests.Text);

            var costumer = new Customer(
                GuestName.Text,
                Nationality.Text,
                RoomNo.Text
                );
            return new InvoiceTemplate(costumer, invoiceDetail, company, rooms, services);

            //var customer = new Customer(GuestName.Text, Nationality.Text,


           //MessageBox.Show(invoiceDate.ToString());
            //return InvoiceTemplate.DefaultTemplate();
        }

        private void generateInvoice(object sender, RoutedEventArgs e)
        {
            try
            {
                var invTemplate = createTemplateForSingleInvoice();
                invTemplate.CreateSingleInvoice();
            }
            catch(Exception ex)
            {
                MessageBox.Show("It looks like there was an error in an entry\n If it looks like all entries are correct, contact developer and give the message below\n" + ex.Message);
            }
            
        }

        private void GenerateBillsFromExcel(object sender, RoutedEventArgs e)
        {
            var startR = int.Parse(StartRow.Text);
            var stopR = int.Parse(EndRow.Text);
            var package = new ExcelPackage(new FileInfo(@"c:\Invoices\GST.xlsx"));
            Dictionary<string, string> addresses = new Dictionary<string, string>
            {
                ["WALK IN"] = " ",
                ["CONCORD"] = "CONCORD EXOTIC VOYAGES(I) PVT.LTD. 407 / 408 - 4TH FLOOR, GERA IMPERIUM II, PATTO, PANJIM, GOA - 403001.",
                ["CAPER"] = "CAPER TRAVEL COMPANY PVT. LTD.HOUSE NO-14/242/B	BEACH PLAZA ANNEXE BUILDING, NEAR KAMAT KINARA, NOMOXIM,CARANZALEM,MIRAMAR, TISWADI, GOA-403002 ",
                ["Minar"] = "Minar Travels (India) Pvt Ltd 101/102, 1st Floor Gera's Imperium II, Patto Plaza Panjim Goa 403001"
            };

            Dictionary<string, string> gst = new Dictionary<string, string>
            {
                ["WALK IN"] = " ",
                ["CONCORD"] = "30AACCC1364LIZZ",
                ["CAPER"] = "30AABCC5600J1Z9",
                ["Minar"] = "30AAACM1267A1ZC"
            };

            ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();
            var start = workSheet.Dimension.Start;
            var end = workSheet.Dimension.End;
            int ctr = 0;
            for (int row = startR; row <= stopR; row++)
            { // Row by row... 
                List<string> tableRow = new List<string>();
                for (int col = start.Column; col <= end.Column; col++)
                { // ... Cell by cell...  
                    var cellValue = workSheet.Cells[row, col].Value.ToString(); // This got me the actual value I needed.  
                    // var cellValue2 = workSheet.Cells[row, col].Value;
                    tableRow.Add(cellValue);
                }
                string invNo = tableRow[0];
                string gName = tableRow[1];
                string agency = tableRow[2];
                string arrival = tableRow[3].Substring(0,10);
                string departure = tableRow[4].Substring(0, 10);
                string noGuests = tableRow[5];
                string lOfStay = tableRow[6];
                string subT = tableRow[7];
                string nationality = tableRow[13];
                string rate = tableRow[10];
                string RMNos = tableRow[9];
                string address = "";
                string gt = "";
                var rmDcrps = RMNos.Split(", ".ToCharArray());
                string invDate = departure;

                gst.TryGetValue(agency, out gt);
                addresses.TryGetValue(agency, out address);
                if(address.Length < 3)
                {
                    address = " ";
                }
                if(gt.Length < 3)
                {
                    gt = " ";
                }

                var k = tableRow[ctr++];
                var k2 = k;
                try
                {
                    var invoiceDetail = new InvoiceDetail(
                departure,
                invNo,
                arrival,
                departure,
                noGuests);

                var costumer = new Customer(
                    gName,
                    nationality,
                    RMNos
                    );
                    var company = new Company(agency, gt, address);
                    List<Room> rooms = new List<Room>();
                    List<Service> services = new List<Service>();
                    foreach(var rm in rmDcrps)
                    {
                        rooms.Add(new Room(rm, double.Parse(rate)));
                    }
                    var invTemplate = new InvoiceTemplate(costumer, invoiceDetail, company, rooms, services);
                    invTemplate.CreateSingleInvoice();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("It looks like there was an error in an entry\n If it looks like all entries are correct, contact developer and give the message below\n" + ex.Message);
                }
            }
        }
    }
}
