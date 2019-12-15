using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using InvoiceDataTemplates;

namespace Invoice
{
    public class InvoiceTemplate
    {
        private const string DefaultTemplateFile = "e:\\Williams Beach Retreat Pvt.docx";

        public string CustomerName { get; set; }
        public string CustomerNationality { get; set; }
        public string RoomNos { get; set; }

        public string InvoiceDate { get; set; }
        public string InvoiceNo { get; set; }
        public string CheckInDate { get; set; }
        public string CheckOutDate { get; set; }
        public int NoOfGuests { get; set; }

        public string CompanyName { get; set; }
        public string CompanyAddress { get; set; }
        public string GSTIN { get; set; }

        public List<Room> Rooms { get; set; }
        public List<Service> Services { get; set; }
        public Company BookingCompany { get; set; }

        public Customer Guest { get; set; }
        public InvoiceDetail InvoiceDetail { get; set; }
        public double Total { get; set; }
        public string TotalInWords { get; set; }

        public InvoiceTemplate(string customerName, string customerNationality, string roomNos,
            string invoiceDate, string invoiceNo, string checkInDate, string checkOutDate, int noOfGuests,
            string companyName, string companyAddress, string gstin)
        {
            CustomerName = customerName;
            CustomerNationality = customerNationality;
            RoomNos = roomNos;

            InvoiceDate = invoiceDate;
            InvoiceNo = invoiceNo;
            CheckInDate = checkInDate;
            CheckOutDate = checkOutDate;
            NoOfGuests = noOfGuests;

            CompanyName = companyName;
            CompanyAddress = companyAddress;
            GSTIN = gstin;
        }

        public InvoiceTemplate(Customer customer, InvoiceDetail detail, Company company, List<Room> rooms, List<Service> services)
        {
            Guest = customer;
            InvoiceDetail = detail;
            BookingCompany = company;
            Rooms = new List<Room>(rooms);
            Services = new List<Service>(services);
            Total = 0;
        }

        public void CreateSingleInvoice(string path = "C:\\Invoices\\InvoiceTemplate.docx")
        {
            Dictionary<string, string> MergeFields = new Dictionary<string, string>
            {
                ["CustomerName"] = Guest.Name,
                ["CustomerNationality"] = Guest.Nationality,
                ["RoomNos"] = Guest.Rooms,
                ["InvoiceDate"] = InvoiceDetail.Date,
                ["InvoiceNo"] = InvoiceDetail.Number,
                ["CheckIn"] = InvoiceDetail.CheckIn,
                ["CheckOut"] = InvoiceDetail.CheckOut,
                ["GuestNo"] = InvoiceDetail.NoGuests,
                ["CompanyName"] = BookingCompany.CompanyName,
                ["CompanyAddress"] = BookingCompany.CompanyAddress,
                ["ComapnyGSTIN"] = BookingCompany.CompanyGST
            };

            var stDate = DateTime.Parse(InvoiceDetail.CheckIn);
            var ebdDt = DateTime.Parse(InvoiceDetail.CheckOut);
            int days = (ebdDt - stDate).Days;

            int rmctr = 1;
            foreach (var rm in Rooms)
            {
                AddRoomDetails(MergeFields, rmctr, rm, days);
                rmctr += 1;
            }

            int svctr = 1;
            foreach (var sv in Services)
            {
                AddSreviceDetails(MergeFields, svctr, sv, days);
                svctr += 1;
            }

            double cleanedTotal = Math.Floor(Total);
            string result = "";
            string totalInWords = GetTotalInWords(cleanedTotal) + " Rupees";
            string cleanTotalText = cleanedTotal.ToString();
            MergeFields["GrandTotalNumber"] = cleanTotalText;
            MergeFields["GrandTotalWords"] = totalInWords;

            TextToWord(path, MergeFields, InvoiceDetail.Number);
        }

        private string GetTotalInWords(double cleanedTotal)
        {
            string total = "";            
            var cleanedTotalWhole = cleanedTotal.ToString();
            var digitToWordSingle = new Dictionary<string, string>
            {
                ["1"] = "One",
                ["2"] = "Two",
                ["3"] = "Three",
                ["4"] = "Four",
                ["5"] = "Five",
                ["6"] = "Six",
                ["7"] = "Seven",
                ["8"] = "Eight",
                ["9"] = "Nine",
            };
            var digitToWordTens = new Dictionary<string, string>
            {
                ["10"] = "Ten",
                ["11"] = "Eleven",
                ["12"] = "Twelve",
                ["13"] = "Thirteen",
                ["14"] = "Fourteen",
                ["15"] = "Fifteen",
                ["16"] = "Sixteen",
                ["17"] = "Seventeen",
                ["18"] = "Eighteen",
                ["19"] = "Nineteen",
                ["0"] = "",
                ["2"] = "Twenty",
                ["3"] = "Thirty",
                ["4"] = "Forty",
                ["5"] = "Fifty",
                ["6"] = "Sixty",
                ["7"] = "Seventy",
                ["8"] = "Eighty",
                ["9"] = "Ninety",
            };
            var isAndUsed = false;
            var hDigit = cleanedTotalWhole.Substring(cleanedTotalWhole.Length - 3, 1);
            if(hDigit == "0")
            {

            }
            else
            {
                total = digitToWordSingle[hDigit] + " Hundred";
            }
            var tANdUDigit = cleanedTotalWhole.Substring(cleanedTotalWhole.Length - 2, 2);
            var tAndUWord = "";
            if (tANdUDigit== "00")
            {

            }
            else
            {
                total += " and " + ResolveTwoDigits(tANdUDigit);
                isAndUsed = true;
            }
            if(cleanedTotalWhole.Length > 3)
            {
                var word = "";
                if (cleanedTotalWhole.Length >= 5)
                {
                    var tDigits = cleanedTotalWhole.Substring(cleanedTotalWhole.Length - 5, 2);
                    word = ResolveTwoDigits(tDigits) + " Thousand";                    
                }
                else
                {
                    var tDigit = cleanedTotalWhole.Substring(cleanedTotalWhole.Length - 4, 1);
                    word = digitToWordSingle[tDigit] + " Thousand";                    
                }
                if (isAndUsed)
                {
                    total = word + ", " + total;
                }
                else
                {
                    if (hDigit != "0")
                    {
                        total = word + " and " + total;
                        isAndUsed = true;
                    }
                    else
                    {
                        total = word + total;
                    }
                }
            }
            if (cleanedTotalWhole.Length > 5)
            {
                var word = "";
                if (cleanedTotalWhole.Length >= 7)
                {
                    var tDigits = cleanedTotalWhole.Substring(cleanedTotalWhole.Length - 7, 2);
                    word = ResolveTwoDigits(tDigits) + " Lakhs";
                }
                else
                {
                    var tDigit = cleanedTotalWhole.Substring(cleanedTotalWhole.Length - 6, 1);
                    word = digitToWordSingle[tDigit] + " Lakhs";
                }
                if (isAndUsed)
                {
                    total = word + ", " + total;
                }
                else
                {
                    if (hDigit != "0")
                    {
                        total = word + " and " + total;
                        isAndUsed = true;
                    }
                    else
                    {
                        total = word + total;
                    }
                }
            }







            //bool firstDigitResolved = false;
            //if (cleanedTotalWhole.Length > 5)
            //{
            //    firstDigitResolved = true;
            //    ResolveLakhs(ref total, ref cleanedTotalWhole, digitToWordSingle, digitToWordTens);
            //}
            //if (cleanedTotalWhole.Length > 3)
            //{
            //    ResolveThousand(ref total, ref cleanedTotalWhole, digitToWordSingle, digitToWordTens, ref firstDigitResolved);
            //}
            //string startSeparator = "";
            //if (firstDigitResolved)
            //{                
            //    var ending = cleanedTotalWhole.Substring(1, 2);
            //    startSeparator = ending == "00" ? " and " : ", ";
            //}

            return total;
        }
        
        private string ResolveTwoDigits(string digits)
        {
            var digitToWordSingle = new Dictionary<string, string>
            {
                ["0"] = "",
                ["1"] = "One",
                ["2"] = "Two",
                ["3"] = "Three",
                ["4"] = "Four",
                ["5"] = "Five",
                ["6"] = "Six",
                ["7"] = "Seven",
                ["8"] = "Eight",
                ["9"] = "Nine",
            };
            var digitToWordTens = new Dictionary<string, string>
            {
                ["10"] = "Ten",
                ["11"] = "Eleven",
                ["12"] = "Twelve",
                ["13"] = "Thirteen",
                ["14"] = "Fourteen",
                ["15"] = "Fifteen",
                ["16"] = "Sixteen",
                ["17"] = "Seventeen",
                ["18"] = "Eighteen",
                ["19"] = "Nineteen",
                ["0"] = "",
                ["2"] = "Twenty ",
                ["3"] = "Thirty ",
                ["4"] = "Forty ",
                ["5"] = "Fifty ",
                ["6"] = "Sixty ",
                ["7"] = "Seventy ",
                ["8"] = "Eighty ",
                ["9"] = "Ninety ",
            };
            string result = "";
            if (digits.StartsWith("1"))
            {
                result = digitToWordTens[digits.Substring(0, 2)];
            }
            else
            {
                result = digitToWordTens[digits.Substring(0, 1)] + digitToWordSingle[digits.Substring(1, 1)];
            }
            return result;
        }

        private static void ResolveThousand(ref string total, ref string cleanedTotalWhole, Dictionary<string, string> digitToWordSingle, Dictionary<string, string> digitToWordTens, ref bool firstDigitResolved)
        {
            if (firstDigitResolved)
            {
                if (cleanedTotalWhole.StartsWith("00"))
                {
                    total += "";
                }
                if (cleanedTotalWhole.StartsWith("0"))
                {
                    total += ", " + digitToWordSingle[cleanedTotalWhole.Substring(1, 1)] + " Thousand";
                    cleanedTotalWhole = cleanedTotalWhole.Substring(2, 3);
                }
                else if (cleanedTotalWhole.StartsWith("1"))
                {
                    total += ", " + digitToWordTens[cleanedTotalWhole.Substring(0, 2)] + " Thousand";
                }
                else
                {
                    total += ", " + digitToWordTens[cleanedTotalWhole.Substring(0, 1)] + " " + digitToWordSingle[cleanedTotalWhole.Substring(1, 1)] + " Thousand";
                }
                cleanedTotalWhole = cleanedTotalWhole.Substring(2, 3);
            }
            else
            {
                firstDigitResolved = true;
                int numberOfDigits = 5;
                if (cleanedTotalWhole.Length == numberOfDigits)
                {
                    if (cleanedTotalWhole.StartsWith("1"))
                    {
                        total += digitToWordTens[cleanedTotalWhole.Substring(0, 2)];
                    }
                    else
                    {
                        total += digitToWordTens[cleanedTotalWhole.Substring(0, 1)] + " " + digitToWordSingle[cleanedTotalWhole.Substring(1, 1)];
                    }
                    cleanedTotalWhole = cleanedTotalWhole.Substring(2, 5);
                }
                else if (cleanedTotalWhole.Length == numberOfDigits - 1)
                {
                    total += digitToWordSingle[cleanedTotalWhole.Substring(0, 1)];
                    cleanedTotalWhole = cleanedTotalWhole.Substring(2, 5);
                }
                total += " Thousand";
            }
        }

        private static void ResolveLakhs(ref string total, ref string cleanedTotalWhole, Dictionary<string, string> digitToWordSingle, Dictionary<string, string> digitToWordTens)
        {
            int numberOfDigits = 7;
            if (cleanedTotalWhole.Length == numberOfDigits)
            {
                if (cleanedTotalWhole.StartsWith("1"))
                {
                    total += digitToWordTens[cleanedTotalWhole.Substring(0, 2)];
                }
                else
                {
                    total += digitToWordTens[cleanedTotalWhole.Substring(0, 1)] + " " + digitToWordSingle[cleanedTotalWhole.Substring(1, 1)];
                }
                cleanedTotalWhole = cleanedTotalWhole.Substring(2, 5);
            }
            else if (cleanedTotalWhole.Length == numberOfDigits - 1)
            {
                total += digitToWordSingle[cleanedTotalWhole.Substring(0, 1)];
                cleanedTotalWhole = cleanedTotalWhole.Substring(2, 5);
            }
            total += " Lakhs";
        }

        private void AddSreviceDetails(Dictionary<string, string> MergeFields, int svctr, Service sv)
        {
            string detailKey = "ServiceDetails" + svctr;
            string rateKey = "ServiceRate" + svctr;
            string cgstKey = "ServiceCGST" + svctr;
            string sgstKey = "ServiceSGST" + svctr;
            string igstKey = "ServiceIGST" + svctr;
            MergeFields[detailKey] = sv.Description;
            MergeFields[rateKey] = sv.Rate.ToString();
            Total += sv.Rate;
            double tax = 0;
            if (sv.Rate < 1000)
            {
                MergeFields[cgstKey] = "0";
                MergeFields[sgstKey] = "0";
                MergeFields[igstKey] = "0";
            }
            else if (sv.Rate < 2500)
            {
                tax = sv.Rate * (12.0 / 100);
                SetGST(MergeFields, cgstKey, sgstKey, igstKey, tax);
            }
            else
            {
                tax = sv.Rate * (18.0 / 100);
                SetGST(MergeFields, cgstKey, sgstKey, igstKey, tax);
            }
            Total += tax;
        }

        private void AddSreviceDetails(Dictionary<string, string> MergeFields, int svctr, Service sv, int days)
        {
            string detailKey = "ServiceDetails" + svctr;
            string rateKey = "ServiceRate" + svctr;
            string cgstKey = "ServiceCGST" + svctr;
            string sgstKey = "ServiceSGST" + svctr;
            string igstKey = "ServiceIGST" + svctr;
            MergeFields[detailKey] = sv.Description;
            MergeFields[rateKey] = sv.Rate.ToString();
            Total += (sv.Rate * days);
            double tax = 0;
            if (sv.Rate < 1000)
            {
                MergeFields[cgstKey] = "0";
                MergeFields[sgstKey] = "0";
                MergeFields[igstKey] = "0";
            }
            else if (sv.Rate < 2500)
            {
                tax = sv.Rate/1.12;
                SetGST(MergeFields, cgstKey, sgstKey, igstKey, tax);
            }
            else
            {
                tax = sv.Rate/1.18;
                SetGST(MergeFields, cgstKey, sgstKey, igstKey, tax);
            }
        }

        private void AddRoomDetails(Dictionary<string, string> MergeFields, int rmctr, Room rm)
        {
            string detailKey = "RoomDetails" + rmctr;
            string rateKey = "RoomRate" + rmctr;
            string cgstKey = "RoomCGST" + rmctr;
            string sgstKey = "RoomSGST" + rmctr;
            string igstKey = "RoomIGST" + rmctr;
            MergeFields[detailKey] = rm.Description;
            MergeFields[rateKey] = rm.Rate.ToString();
            Total += rm.Rate;
            double tax = 0;
            if (rm.Rate < 1000)
            {
                MergeFields[cgstKey] = "0";
                MergeFields[sgstKey] = "0";
                MergeFields[igstKey] = "0";
            }
            else if (rm.Rate < 2500)
            {
                tax = rm.Rate * (12.0 / 100);
                SetGST(MergeFields, cgstKey, sgstKey, igstKey, tax);
            }
            else
            {
                tax = rm.Rate * (18.0 / 100);
                SetGST(MergeFields, cgstKey, sgstKey, igstKey, tax);
            }
            Total += tax;
        }

        private void AddRoomDetails(Dictionary<string, string> MergeFields, int rmctr, Room rm, int days)
        {
            string detailKey = "RoomDetails" + rmctr;
            string rateKey = "RoomRate" + rmctr;
            string cgstKey = "RoomCGST" + rmctr;
            string sgstKey = "RoomSGST" + rmctr;
            string igstKey = "RoomIGST" + rmctr;
            MergeFields[detailKey] = rm.Description;
            MergeFields[rateKey] = rm.Rate.ToString();
            Total += (rm.Rate * days);
            double tax = 0;
            if (rm.Rate < 1000)
            {
                MergeFields[cgstKey] = "0";
                MergeFields[sgstKey] = "0";
                MergeFields[igstKey] = "0";
            }
            else if (rm.Rate < 2500)
            {
                tax = rm.Rate / 1.12;
                SetGST(MergeFields, cgstKey, sgstKey, igstKey, tax);
            }
            else
            {
                tax = rm.Rate/1.18 ;
                SetGST(MergeFields, cgstKey, sgstKey, igstKey, tax);
            }
        }

        private void SetGST(Dictionary<string, string> MergeFields, string cgstKey, string sgstKey, string igstKey, double tax)
        {
            if (BookingCompany.CompanyGST.StartsWith("30"))
            {
                MergeFields[cgstKey] = Math.Round(tax / 2, 2).ToString();
                MergeFields[sgstKey] = Math.Round(tax / 2, 2).ToString();
                MergeFields[igstKey] = "0";

            }
            else
            {
                MergeFields[cgstKey] = "0";
                MergeFields[sgstKey] = "0";
                MergeFields[igstKey] = tax.ToString();
            }
        }

        public void ModifyDocument()
        {
            Dictionary<string, string> MergeFields = CreateMergeDictionary();
            string path = DefaultTemplateFile;
            TextToWord(path, MergeFields);
        }

        private Dictionary<string, string> CreateMergeDictionary()
        {
            var mergeList =  new Dictionary<string, string>
            {
                ["CustomerName"] = CustomerName,
                ["CustomerNationality"] = CustomerNationality,
                ["RoomNos"] = RoomNos,
                ["InvoiceDate"] = InvoiceDate,
                ["InvoiceNo"] = InvoiceNo,
                ["CheckIn"] = CheckInDate,
                ["CheckOut"] = CheckOutDate,
                ["GuestNo"] = NoOfGuests.ToString(),
                ["CompanyName"] = CompanyName,
                ["CompanyAddress"] = CompanyAddress,
                ["ComapnyGSTIN"] = GSTIN
            };

            return mergeList;
        }

        public static void TextToWord(string pWordDoc, Dictionary<string, string> pDictionaryMerge)
        {
            Object oMissing = System.Reflection.Missing.Value;
            Object oTrue = true;
            Object oFalse = false;
            Application oWord = new Application();
            Document oWordDoc = new Document();
            oWord.Visible = true;
            Object oTemplatePath = pWordDoc;
            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            List<string> mergeFields = new List<string>();

            foreach (Microsoft.Office.Interop.Word.Field myMergeField in oWordDoc.Fields)
            {
                Microsoft.Office.Interop.Word.Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;
                if (fieldText.StartsWith(" MERGEFIELD"))
                {
                    Int32 endMerge = fieldText.IndexOf("\\");
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(11, endMerge - 11);
                    fieldName = fieldName.Trim();
                    mergeFields.Add(fieldName);
                    string outText = "";
                    if(pDictionaryMerge.TryGetValue(fieldName, out outText))
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(outText);
                    }
                    else
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(" ");
                    }
                }
            }

            //using (TextWriter tw = new StreamWriter("e:\\SavedList2.txt"))
            //{
            //    foreach (string s in mergeFields)
            //        tw.WriteLine(s);
            //}
            if (!Directory.Exists(@"C:\test")){
                Directory.CreateDirectory(@"C:\test");
            }
                

            Object oSaveAsFile = (Object)"e:\\NewTestFolder\\SampleDocajmal11.docx";
            oWordDoc.SaveAs(ref oSaveAsFile, ref oMissing, ref oMissing, ref oMissing,

                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,

                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,

                ref oMissing, ref oMissing);

            oWordDoc.ExportAsFixedFormat("e:\\NewTestFolder\\SampleDocajmal11.pdf", WdExportFormat.wdExportFormatPDF);

            oWordDoc.Close(ref oFalse, ref oMissing, ref oMissing);

            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);

        }

        public static void TextToWord(string pWordDoc, Dictionary<string, string> pDictionaryMerge, string filename)
        {
            Object oMissing = System.Reflection.Missing.Value;
            Object oTrue = true;
            Object oFalse = false;
            Application oWord = new Application();
            Document oWordDoc = new Document();
            oWord.Visible = true;
            Object oTemplatePath = pWordDoc;
            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            List<string> mergeFields = new List<string>();

            foreach (Microsoft.Office.Interop.Word.Field myMergeField in oWordDoc.Fields)
            {
                Microsoft.Office.Interop.Word.Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;
                if (fieldText.StartsWith(" MERGEFIELD"))
                {
                    Int32 endMerge = fieldText.IndexOf("\\");
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(11, endMerge - 11);
                    fieldName = fieldName.Trim();
                    mergeFields.Add(fieldName);
                    string outText = "";
                    if (pDictionaryMerge.TryGetValue(fieldName, out outText))
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(outText);
                    }
                    else
                    {
                        myMergeField.Select();
                        oWord.Selection.TypeText(" ");
                    }
                }
            }

            //using (TextWriter tw = new StreamWriter("e:\\SavedList2.txt"))
            //{
            //    foreach (string s in mergeFields)
            //        tw.WriteLine(s);
            //}
            if (!Directory.Exists(@"C:\Invoices"))
            {
                Directory.CreateDirectory(@"C:\Invoices");
            }

            var path = Path.Combine(@"C:\Invoices", filename + ".docx");

            Object oSaveAsFile = (Object)path;
            oWordDoc.SaveAs(ref oSaveAsFile, ref oMissing, ref oMissing, ref oMissing,

                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,

                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,

                ref oMissing, ref oMissing);
            var path2 = Path.Combine(@"C:\Invoices", filename + ".pdf");

            oWordDoc.ExportAsFixedFormat(path2, WdExportFormat.wdExportFormatPDF);

            oWordDoc.Close(ref oFalse, ref oMissing, ref oMissing);

            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);

        }

        public static InvoiceTemplate DefaultTemplate()
        {
            return new InvoiceTemplate("Ajmal Sohail Kukkangai 1", "Indian", "109, C3, C4",
                "02-02-2020", "24ISAF", "25-01-2020", "02-02-2020", 6,
                "Booking.com", "RMZ Millennia, No.1 & 2, Murphy Road, Tower B – 6th floor, Ulsoor, Bangalore, 560008, India", "42AAACW4430F1ZX");
        }
    }
}
