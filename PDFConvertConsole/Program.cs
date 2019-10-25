using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PDFConvertConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Customer> customers = new List<Customer>();
            List<String> outputLines = new List<String>();

            //C:\Users\michael.march\Downloads\projections.csv
            //String openSaleFile = "C: \\Users\\michael.march\\Documents\\test\\BHFI2\\091718_BHFinancial_$5.18mm_LoanMe_Open_Sale_File.csv";
            //String baseFolder = "C: \\Users\\michael.march\\Documents\\test\\loanmetest\\test2";
            String baseFolder = @"M:\Loan Me Media\Contracts";
            //String baseFolder = @"C: \Users\michael.march\Documents\test\loanmetest\loanme_20190708";

               //String remainingAccountsLocation = baseFolder + "\\outputAccountsTake2excel.xlsx";
               //String remainingAccountsLocation = txtAccountLocation.Text;

               //Excel.Application xlApp1 = new Excel.Application();
               //Excel.Workbook xlWorkbook1 = xlApp1.Workbooks.Open(remainingAccountsLocation);
               //Excel._Worksheet xlWorksheet1 = xlWorkbook1.Sheets[1];
               //Excel.Range xlRange1 = xlWorksheet1.UsedRange;

               //int rowCount1 = xlRange1.Rows.Count;
               //string[] remainingAccount = new string[rowCount1];
               ////iterate over the rows and columns and print to the console as it appears in the file
               ////excel is not zero based!!
               //for (int i = 2; i <= rowCount1; i++)
               //{
               //    if (xlRange1.Cells[i, 3] != null && xlRange1.Cells[i, 3].Value2 != null)
               //    {
               //        //txtScreen.Text = txtScreen.Text + xlRange1.Cells[i, 3].Value2 + "\n";
               //        remainingAccount[i-2] = Convert.ToString(xlRange1.Cells[i, 3].Value2);
               //    }
               //}

               ////release com objects to fully kill excel process from running in the background
               //Marshal.ReleaseComObject(xlRange1);
               //Marshal.ReleaseComObject(xlWorksheet1);

               ////close and release
               //xlWorkbook1.Close();
               //Marshal.ReleaseComObject(xlWorkbook1);

               ////quit and release
               //xlApp1.Quit();
               //Marshal.ReleaseComObject(xlApp1);

            String openSaleFile = baseFolder + "\\LoanMeDataCSV.csv";
            StreamReader reader = new StreamReader(openSaleFile);

            //txtScreen.Text = "remaining accounts: " + remainingAccount.Length;

            int linecount = 0;
            while (!reader.EndOfStream)
            {
                linecount++;
                string line = reader.ReadLine();
                if (!String.IsNullOrWhiteSpace(line) && linecount > 1)
                {
                    string[] values = line.Split(',');
                    //Boolean idFound = false;
                    //for (int i = 0; i < remainingAccount.Length; i++)
                    //{
                    //    if (remainingAccount[i] == values[0].Replace("\"", ""))
                    //    {
                    //        txtScreen.Text = txtScreen.Text + remainingAccount[i] + " found\n";
                    //        idFound = true;
                    //    }
                    //}
                    if (values.Length >= 4 && values[4].Length > 0)
                    {
                        Customer customer = new Customer();
                        //customer.fileNo = values[0].Replace("\"", "");
                        customer.loanid = values[0].Replace("\"", "");
                        //customer.lastName = values[5];
                        //customer.firstName = values[6];
                        customer.fullName = values[1].Replace("\"", "") + " " + values[2].Replace("\"", "");
                        customer.interestRate = values[5].Replace("\"", "");
                        //customer.zip = values[11];
                        customer.principalBalance = values[9].Replace("\"", "");
                        customer.dmpPayments = values[8].Replace("\"", "");
                        customer.originationDate = values[3].Replace("\"", "");
                        customer.chargeOffDate = values[4].Replace("\"", "");
                        customer.chargeOffInterest = values[6].Replace("\"", "");
                        customer.chargeOffBalance = values[7].Replace("\"", "");
                        customers.Add(customer);
                        //txtScreen.Text = txtScreen.Text + customer.print() + "\n";
                    }
                }

            }
            reader.Close();

            Console.WriteLine( DateTime.Now + " Customers: " + customers.Count);

            ///////////////////////////////////
            // Get folders                   //
            ///////////////////////////////////
            string folderLocal = baseFolder;
            string[] fileNames = Directory.GetDirectories(folderLocal, "*.*", SearchOption.AllDirectories);
            int countblah = 1;
            int customersProcessed = 1;
            foreach (string fileName in fileNames)
            {
                if (!fileName.Contains("Newfolder"))
                {
                    Customer thisCustomer = null;
                    string tempLoanId = "N/A";
                    try
                    {
                        string tempFileName = fileName.Substring(folderLocal.Length + 1);
                        tempLoanId = tempFileName.Substring(0, tempFileName.IndexOf(" "));
                        string tempString = tempFileName.Substring(tempFileName.IndexOf(" ") + 1);
                        string tempFileNo = tempString.Substring(0, tempString.IndexOf(" ") + 1);
                      
                        foreach (Customer customer in customers)
                        {
                            if (customer.loanid == tempLoanId)
                            {
                                customer.fileNo = tempFileNo;
                                //customer.interestRate = getInterestRate(tempFileNo);   // No longer needed to get the interest rate I think
                                thisCustomer = customer;
                                customersProcessed++;
                                Console.WriteLine(customersProcessed + "/" + customers.Count);
                                break;
                            }
                        }
                    }
                    catch (System.ArgumentOutOfRangeException) {
                        Console.WriteLine("Improper folder format");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(string.Format(ex.ToString()));
                    }
                 
                    Console.WriteLine("Files parsed in drive: " + countblah.ToString());
                    countblah++;
                    if (thisCustomer != null)
                    {
                        Console.WriteLine("Processing file: " + fileName.Substring(folderLocal.Length + 1) + " " + tempLoanId);
                        String loadFile = fileName + "\\TRANS_" + thisCustomer.loanid + ".xls";
                        try
                        {
                            //reader = new StreamReader(loadFile);
                            //linecount = 0;
                            //while (!reader.EndOfStream)
                            //{
                            //    string line = reader.ReadLine();
                            //    linecount++;
                            //    if (!String.IsNullOrWhiteSpace(line) && linecount > 1)
                            //    {
                            //        string[] values = line.Split(',');
                            //        if (values.Length >= 4 && values[4].Length > 3)
                            //        {
                            //            //txtScreen.Text = txtScreen.Text + values[0] + " " + values[2] + " " + values[3] + "\n";
                            //            String outputString = linecount - 1 + ",\"" + thisCustomer.fileNo + "\"," + "\"" + thisCustomer.fullName + "\"," + "\"" + thisCustomer.principalBalance + "\"," + "\"" + thisCustomer.interestRate + "\"," + line;
                            //            outputLines.Add(outputString);
                            //        }
                            //    }
                            //}
                            //reader.Close();
                            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            //Create COM Objects. Create a COM object for everything that is referenced
                            Excel.Application xlApp = new Excel.Application();
                            Console.Write(loadFile);
                            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(loadFile);
                            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                            Excel.Range xlRange = xlWorksheet.UsedRange;
                            int rowCount = xlRange.Rows.Count;
                            int colCount = xlRange.Columns.Count;

                            //iterate over the rows and columns and print to the console as it appears in the file
                            //excel is not zero based!!
                            for (int i = 2; i <= rowCount; i++)
                            {
                                //for (int j = 1; j <= colCount; j++)
                                //{
                                //new line
                                //if (j == 1)                      

                                //write the value to the console
                                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                                {
                                    String lineout = xlRange.Cells[i, 1].Value2.ToString();

                                    if (xlRange.Cells[i, 2].Value2 != null)
                                    {
                                        lineout = lineout + "," + xlRange.Cells[i, 2].Value2;
                                    }
                                    else
                                    {
                                        lineout = lineout + ",";
                                    }
                                    if (xlRange.Cells[i, 3].Value2 != null)
                                    {
                                        lineout = lineout + "," + DateTime.FromOADate(xlRange.Cells[i, 3].Value2).ToString("yyyy-MM-dd");
                                    }
                                    else
                                    {
                                        lineout = lineout + ",";
                                    }
                                    if (xlRange.Cells[i, 4].Value2 != null)
                                    {
                                        lineout = lineout + "," + xlRange.Cells[i, 4].Value2;
                                    }
                                    else
                                    {
                                        lineout = lineout + ",";
                                    }
                                    if (xlRange.Cells[i, 5].Value2 != null)
                                    {
                                        lineout = lineout + "," + xlRange.Cells[i, 5].Value2;
                                    }
                                    else
                                    {
                                        lineout = lineout + ",";
                                    }
                                    if (xlRange.Cells[i, 6].Value2 != null)
                                    {
                                        lineout = lineout + "," + xlRange.Cells[i, 6].Value2;
                                    }
                                    else
                                    {
                                        lineout = lineout + ",";
                                    }
                                    if (xlRange.Cells[i, 7].Value2 != null)
                                    {
                                        lineout = lineout + "," + xlRange.Cells[i, 7].Value2;
                                    }
                                    else
                                    {
                                        lineout = lineout + ",";
                                    }
                                    if (xlRange.Cells[i, 8].Value2 != null)
                                    {
                                        lineout = lineout + "," + xlRange.Cells[i, 8].Value2;
                                    }
                                    else
                                    {
                                        lineout = lineout + ",";
                                    }
                                    if (xlRange.Cells[i, 9].Value2 != null)
                                    {
                                        lineout = lineout + "," + xlRange.Cells[i, 9].Value2;
                                    }
                                    else
                                    {
                                        lineout = lineout + ",";
                                    }
                                    if (xlRange.Cells[i, 10].Value2 != null)
                                    {
                                        lineout = lineout + "," + xlRange.Cells[i, 10].Value2;
                                    }
                                    else
                                    {
                                        lineout = lineout + ",";
                                    }
                                    //String lineout = loanid + "," + tno + "," + newDate + "," + tcode + "," + transtype + "," + paymentamount + "," + principal + "," + interest + "," + fees + "," + principalbalance;
                                    //txtScreen.Text = txtScreen.Text + lineout + "\n";
                                    String outputString = i - 1 + ",\"" + thisCustomer.fileNo + "\"," + "\"" + thisCustomer.fullName + "\"," + "\"" + thisCustomer.principalBalance + "\"," + "\"" + thisCustomer.interestRate + "\"," + "\"" + thisCustomer.originationDate + "\"," + "\"" + thisCustomer.chargeOffDate + "\"," + "\"" + thisCustomer.chargeOffInterest + "\"," + "\"" + thisCustomer.chargeOffBalance + "\"," + "\"" + thisCustomer.dmpPayments + "\"," + lineout;
                                    outputLines.Add(outputString);
                                }
                            }
                            //release com objects to fully kill excel process from running in the background
                            Marshal.ReleaseComObject(xlRange);
                            Marshal.ReleaseComObject(xlWorksheet);

                            //close and release
                            xlWorkbook.Close();
                            Marshal.ReleaseComObject(xlWorkbook);

                            //quit and release
                            xlApp.Quit();
                            Marshal.ReleaseComObject(xlApp);
                            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(string.Format(ex.ToString()));
                        }
                    }
                    else
                    {
                        Console.WriteLine(tempLoanId + " not found");
                    }
                }
            }
            //////////////////////////////////////////
            // End get folders                      //
            //////////////////////////////////////////
            ///
            //foreach (Customer customer in customers)
            //{
            //    //txtScreen.Text = txtScreen.Text + customer.print() + "\n";
            //    String loadFile = baseFolder + "O:\\Misc\\miketest\\mikeinput\\" + customer.loanid + " " + customer.fileNo + " CA BHFI\\TRANS_" + customer.loanid + ".csv";
            //    try
            //    {
            //        reader = new StreamReader(loadFile);
            //        linecount = 0;
            //        while (!reader.EndOfStream)
            //        {
            //            string line = reader.ReadLine();
            //            linecount++;
            //            if (!String.IsNullOrWhiteSpace(line) && linecount > 1)
            //            {
            //                string[] values = line.Split(',');
            //                if (values.Length >= 4 && values[4].Length > 3)
            //                {
            //                    //txtScreen.Text = txtScreen.Text + values[0] + " " + values[2] + " " + values[3] + "\n";
            //                    String outputString = linecount - 1 + ",\"" + customer.fileNo + "\"," + "\"" + customer.fullName + "\"," + "\"" + customer.principalBalance + "\"," + "\"" + customer.interestRate + "\"," + line;
            //                    outputLines.Add(outputString);
            //                }
            //            }
            //        }
            //        reader.Close();
            //    }
            //    catch (System.IO.DirectoryNotFoundException)
            //    {
            //        txtScreen.Text = txtScreen.Text + customer.loanid + " folder was not found\n";
            //    }
            //    catch (System.IO.FileNotFoundException)
            //    {
            //        txtScreen.Text = txtScreen.Text + customer.loanid + " file was not found\n";
            //    }
            //}
            Console.WriteLine("Writing Output");
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(baseFolder + "\\outputTest.csv"))
            {
                String firstLine = "ID,Number,FILE_NO,FULL_NAME,PRINCIPAL,INTEREST_RATE,ORIGDATE,CHARGEOFFDATE,CHARGEOFFINTEREST,COBALANCE,DMPPAYMENTS,LOAN_ID,T_NO,APPLY_DATE,T_CODE,TRANS_TYPE,PAYMENT_AMOUNT,PRINCIPAL,INTEREST,FEES,Field15";
                file.WriteLine(firstLine);
                int count = 1;
                foreach (String outputLine in outputLines)
                {
                    file.WriteLine(count + "," + outputLine);
                    count++;
                }
                Console.WriteLine("OutputCount: " + count);
            }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(baseFolder + "\\outputAccounts.csv"))
            {
                String firstLine = "ID,FILE_NO,LOAN_ID";
                file.WriteLine(firstLine);
                int count = 1;
                foreach (Customer customer in customers)
                {
                    if (customer.fileNo != "0")
                    {
                        file.WriteLine(count + "," + customer.fileNo + "," + customer.loanid);
                        count++;
                    }
                }
            }

            Console.WriteLine("CustomerCount: " + customers.Count());
            Console.WriteLine("Press Enter Twice to Exit");
            Console.ReadLine();
        }
    }
}
