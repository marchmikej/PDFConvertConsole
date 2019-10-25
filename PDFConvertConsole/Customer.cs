using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFConvertConsole
{
    class Customer
    {
        public String fileNo = "0";
        public String loanid = "";
        public String lastName = "";
        public String firstName = "";
        public String fullName = "";
        public String zip = "";
        public String principalBalance = "";
        public String interestRate = "";
        public String dmpPayments = "";
        public String originationDate = "";
        public String chargeOffDate = "";
        public String chargeOffInterest = "";
        public String chargeOffBalance = "";

        public String print()
        {
            return this.fileNo + "," + this.loanid + "," + this.fullName + "," + this.interestRate + "," + this.zip + "," + this.principalBalance + "," + this.dmpPayments + "," + this.originationDate + "," + chargeOffDate + "," + "," + chargeOffInterest + "," + chargeOffBalance;
        }
    }
}
