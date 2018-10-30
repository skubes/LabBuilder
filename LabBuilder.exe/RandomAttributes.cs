using System;
using System.Collections.Generic;
using System.Text;


namespace LabBuilder
{
    class RandomAttributes
    {


        public RandomAttributes()
        {
            // initialize fields with random data;

            Organization = attributes["orgs"][rg.Next(attributes["orgs"].Length)];
            Title = attributes["titles"][rg.Next(attributes["titles"].Length)];
            Country = attributes["countries"][rg.Next(attributes["countries"].Length)];
            Department = attributes["depts"][rg.Next(attributes["depts"].Length)];
            EmployeeID = attributes["empolyeeids"][rg.Next(attributes["empolyeeids"].Length)];
            Office = attributes["offices"][rg.Next(attributes["offices"].Length)];
            HomePhone = attributes["homephones"][rg.Next(attributes["homephones"].Length)];
            OfficePhone = attributes["officephones"][rg.Next(attributes["officephones"].Length)];
            PostalCode = attributes["postals"][rg.Next(attributes["postals"].Length)];
            State = attributes["states"][rg.Next(attributes["states"].Length)];
            City = attributes["cities"][rg.Next(attributes["cities"].Length)];
            StreetAddress = attributes["streets"][rg.Next(attributes["streets"].Length)];
            Division = attributes["divs"][rg.Next(attributes["divs"].Length)];


        }

        public string Title;
        public string Country;
        public string Department;
        public string EmployeeID;
        public string Office;
        public string HomePhone;
        public string OfficePhone;
        public string PostalCode;
        public string State;
        public string City;
        public string StreetAddress;
        public string Division;
        public string Organization;

        static Random rg = new Random();

        static Dictionary<string, string[]> attributes = PopulateRandomAttributes();

        static Dictionary<string, string[]> PopulateRandomAttributes()
        {
            var atts = new Dictionary<string, string[]>();
            atts.Add("orgs", new string[] { "Information Availability", "Information Intelligence", "Backup and Recovery" });
            atts.Add("divs", new string[] { "EMEA", "Americas", "Asia Pacific" });
            atts.Add("streets", new string[] { "11 Guardians Way", "1620 Columbus Cir", "400 Rude Awakening Blvd", "1 Software Loop", "4 Garbage Collection Circle" });
            atts.Add("cities", new string[] { "Lake Mary", "New York", "Las Vegas", "Enfield", "Los Alamos" });
            atts.Add("states", new string[] { "Florida", "New York", "Nevada", "Connecticut", "New Mexico" });
            atts.Add("postals", new string[] { "32746", "20102", "90210", "06082", "28375" });
            atts.Add("officephones", new string[] { "321-746-2384", "+44 2356 358 85383", "533-238-2586", "101-420-5209", "284-288-4284" });
            atts.Add("homephones", new string[] { "353-583-39385", "253-385-3959", "217-413-3253", "567-352-5353", "234-099-3888" });
            atts.Add("offices", new string[] { "Heathrow, Florida", "Las Vegas, Nevada", "Enfield, Connecticut", "Los Alamos, New Mexico" });
            atts.Add("empolyeeids", new string[] { "4532746", "220102", "990210", "0006082", "258375" });
            atts.Add("depts", new string[] { "Technical Support", "Information Technology", "Marketing", "Legal", "Products" });
            atts.Add("countries", new string[] { "US", "GB", "CA" });
            atts.Add("titles", new string[] { "Software Engineer", "Sr Software Engineer", "Princ Software Engineer", "Sr Princ Software Engineer", "Assoc Software Engineer", "Staff Software Engineer", "Program Manager", "Director of Development", "QA Engineer", "Tech Support Engineer", "Sr Tech Support Engineer", "Princ Tech Support Engineer", "Sr Princ Tech Support Engineer", "Manger", "Sr Manger", "Scrum Master", "Product Manager", "Vice President of Sales" });
            return atts;
        }
    }
}
