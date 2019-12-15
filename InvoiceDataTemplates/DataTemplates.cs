using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceDataTemplates
{
    public class DataTemplates
    {
    }

    public struct Room
    {
        public string Description;
        public double Rate;

        public Room(string desc, double rate)
        {
            Description = desc;
            Rate = rate;
        }
    }

    public struct InvoiceDetail
    {
        public string Date;
        public string Number;
        public string CheckIn;
        public string CheckOut;
        public string NoGuests;

        public InvoiceDetail(string date, string number, string checkIn, string checkOut, string guests)
        {
            Date = date;
            Number = number;
            CheckIn = checkIn;
            CheckOut = checkOut;
            NoGuests = guests;
        }
    }

    public struct Customer
    {
        public string Name;
        public string Nationality;
        public string Rooms;

        public Customer(string name, string nation, string rooms)
        {
            Name = name;
            Nationality = nation;
            Rooms = rooms;
        }
    }

    public struct Service
    {
        public string Description;
        public double Rate;

        public Service(string desc, double rate)
        {
            Description = desc;
            Rate = rate;
        }
    }

    public struct Company
    {
        public string CompanyName;
        public string CompanyGST;
        public string CompanyAddress;        

        public Company(string companyName)
        {
            Dictionary<string, string> addresses = new Dictionary<string, string>
            {
                ["WALK IN"] = " ",
                ["BOOKING DOT COM"] = "WQ9012",
                ["GOIBIBO"] = "sadfgqa ",
                ["MAKE MY TRIP"] = "safd edsayhrt",
                ["DESIYA"] = "asedf afg"
            };

            Dictionary<string, string> gst = new Dictionary<string, string>
            {
                ["WALK IN"] = " ",
                ["BOOKING DOT COM"] = "301132421",
                ["GOIBIBO"] = "3012412",
                ["MAKE MY TRIP"] = "2141353",
                ["DESIYA"] = "124531563"
            };

            CompanyName = companyName;
            CompanyGST = gst[companyName];
            CompanyAddress = addresses[companyName];            
        }
        public Company(string companyName, string gst, string address)
        {
            CompanyName = companyName;
            CompanyGST = gst;
            CompanyAddress = address;
        }
    }
}
