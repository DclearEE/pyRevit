using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NBK.Spool
{
    public class SpoolNumber
    {
        public string FullSpoolNo { get; set; }
        public int Number { get; set; }
        public int ArrayCount { get; set; }
        public string Abbreviation1 { get; set; }
        public string Abbreviation2 { get; set; }
        public string FloorCode { get; set; }
        public string SystemAbbreviation { get; set; }
        public string FamilySizeandLength { get; set; }
        public string CategoryName { get; set; }
        public string PartType { get; set; }


        public SpoolNumber()
        {
        }

        public SpoolNumber(string fullSpoolNo)
        {
            this.FullSpoolNo = fullSpoolNo;

            ParseString(fullSpoolNo);
        }

        private void ParseString(string fullSpoolNo)
        {
            // Split the spool numbers
            string[] SpoolNoArray = fullSpoolNo.Split('-');

            this.ArrayCount = SpoolNoArray.Count();

            // Parse the number
            int number = 0;
            Int32.TryParse(SpoolNoArray[this.ArrayCount - 1], out number);
            this.Number = number;

            if(this.ArrayCount > 1)
            {
                this.Abbreviation1 = SpoolNoArray[0].Trim();
            }

            if (this.ArrayCount > 2)
            {
                this.Abbreviation2 = SpoolNoArray[1].Trim();
            }
        }
    }
}
