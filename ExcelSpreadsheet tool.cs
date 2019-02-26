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
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;


namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // Daily check box to check and uncheck reports

        private void dailyCheck_CheckedChanged(object sender, RoutedEventArgs e)
        {
            bool newVal = (dailyCheck.IsChecked == true);
            
            AstoriaLpFldRenewal5box.IsChecked = newVal;
            Bridgebox.IsChecked = newVal;
            CancellationNonPayBox.IsChecked = newVal;
            DastoriaLPBox.IsChecked = newVal;
            DisbStopbox.IsChecked = newVal;
            LoansNoInsbox.IsChecked = newVal;
            lpFldsfhaBox.IsChecked = newVal;
            invalidPayeebox.IsChecked = newVal;
            mgcnoBox.IsChecked = newVal;
            mgcnoNotebox.IsChecked = newVal;
            notRolledbox.IsChecked = newVal;
            rolledFarbox.IsChecked = newVal;
            Cliqbebox.IsChecked = newVal;
            Hazanabox.IsChecked = newVal;
            Hazdisbox.IsChecked = newVal;
            Hazfcobox.IsChecked = newVal;
            Hazfhabox.IsChecked = newVal;
            Hazfo1box.IsChecked = newVal;
            Hazlegbox.IsChecked = newVal;
            Hazpoibox.IsChecked = newVal;
            Hazsalbox.IsChecked = newVal;
            Inschgbox.IsChecked = newVal;
            Insuptbox.IsChecked = newVal;
            Prtqbebox.IsChecked = newVal;
            Qbeblkbox.IsChecked = newVal;
            Qbepoibox.IsChecked = newVal;
            Reoaddbox.IsChecked = newVal;
            Reocnxbox.IsChecked = newVal;
            Respipbox.IsChecked = newVal;
            Zcirptbox.IsChecked = newVal;
            Zcscrubox.IsChecked = newVal;
            Zcscssbox.IsChecked = newVal;
            Zcsdocbox.IsChecked = newVal;
            Zcshotbox.IsChecked = newVal;
            Zcslp1box.IsChecked = newVal;
            Zcsoutbox.IsChecked = newVal;
            Flddecbox.IsChecked = newVal;
            Flddf1box.IsChecked = newVal;
            Flddf2box.IsChecked = newVal;
            Fldec1box.IsChecked = newVal;
            Fldec2box.IsChecked = newVal;
            Fldzdrbox.IsChecked = newVal;
            Lpnot1box.IsChecked = newVal;
            Lpnot2box.IsChecked = newVal;
            Lpnot3box.IsChecked = newVal;


        }

        private void dailyBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            dailyCheck.IsChecked = null;
            if ((AstoriaLpFldRenewal5box.IsChecked == true) && (Bridgebox.IsChecked == true) && (CancellationNonPayBox.IsChecked == true) && (DastoriaLPBox.IsChecked == true) && (DisbStopbox.IsChecked == true) && (LoansNoInsbox.IsChecked == true) && (lpFldsfhaBox.IsChecked == true) && (invalidPayeebox.IsChecked == true) && (mgcnoBox.IsChecked == true) && (mgcnoNotebox.IsChecked == true) && (notRolledbox.IsChecked == true) && (rolledFarbox.IsChecked == true) && (Cliqbebox.IsChecked == true) && (Hazanabox.IsChecked == true) && (Hazdisbox.IsChecked == true) && (Hazfcobox.IsChecked == true) && (Hazfhabox.IsChecked == true) && (Hazfo1box.IsChecked == true) && (Hazlegbox.IsChecked == true) && (Hazpoibox.IsChecked == true) && (Hazsalbox.IsChecked == true) && (Inschgbox.IsChecked == true) && (Insuptbox.IsChecked == true) && (Prtqbebox.IsChecked == true) && (Qbeblkbox.IsChecked == true) && (Qbepoibox.IsChecked == true) && (Reoaddbox.IsChecked == true) && (Reocnxbox.IsChecked == true) && (Respipbox.IsChecked == true) && (Zcirptbox.IsChecked == true) && (Zcscrubox.IsChecked == true) && (Zcscssbox.IsChecked == true) && (Zcsdocbox.IsChecked == true) && (Zcshotbox.IsChecked == true) && (Zcslp1box.IsChecked == true) && (Zcsoutbox.IsChecked == true) && (Flddecbox.IsChecked == true) && (Flddf1box.IsChecked == true) && (Flddf2box.IsChecked == true) && (Fldec1box.IsChecked == true) && (Fldec2box.IsChecked == true) && (Fldzdrbox.IsChecked == true) && (Lpnot1box.IsChecked == true) && (Lpnot2box.IsChecked == true) && (Lpnot3box.IsChecked == true))
                dailyCheck.IsChecked = true;
            if ((AstoriaLpFldRenewal5box.IsChecked == false) && (Bridgebox.IsChecked == false) && (CancellationNonPayBox.IsChecked == false) && (DastoriaLPBox.IsChecked == false) && (DisbStopbox.IsChecked == false) && (LoansNoInsbox.IsChecked == false) && (lpFldsfhaBox.IsChecked == false) && (invalidPayeebox.IsChecked == false) && (mgcnoBox.IsChecked == false) && (mgcnoNotebox.IsChecked == false) && (notRolledbox.IsChecked == false) && (rolledFarbox.IsChecked == false) && (Cliqbebox.IsChecked == false) && (Hazanabox.IsChecked == false) && (Hazdisbox.IsChecked == false) && (Hazfcobox.IsChecked == false) && (Hazfhabox.IsChecked == false) && (Hazfo1box.IsChecked == false) && (Hazlegbox.IsChecked == false) && (Hazpoibox.IsChecked == false) && (Hazsalbox.IsChecked == false) && (Inschgbox.IsChecked == false) && (Insuptbox.IsChecked == false) && (Prtqbebox.IsChecked == false) && (Qbeblkbox.IsChecked == false) && (Qbepoibox.IsChecked == false) && (Reoaddbox.IsChecked == false) && (Reocnxbox.IsChecked == false) && (Respipbox.IsChecked == false) && (Zcirptbox.IsChecked == false) && (Zcscrubox.IsChecked == false) && (Zcscssbox.IsChecked == false) && (Zcsdocbox.IsChecked == false) && (Zcshotbox.IsChecked == false) && (Zcslp1box.IsChecked == false) && (Zcsoutbox.IsChecked == false) && (Flddecbox.IsChecked == false) && (Flddf1box.IsChecked == false) && (Flddf2box.IsChecked == false) && (Fldec1box.IsChecked == false) && (Fldec2box.IsChecked == false) && (Fldzdrbox.IsChecked == false) && (Lpnot1box.IsChecked == false) && (Lpnot2box.IsChecked == false) && (Lpnot3box.IsChecked == false))
                dailyCheck.IsChecked = false;

        }

        // Weekly reports checkbox to check and uncheck

        private void wkCheck_CheckedChanged(object sender, RoutedEventArgs e)
        {
            bool newFal = (wkCheck.IsChecked == true);
            HeaderFileHaBox.IsChecked = newFal;
            ExpLineBox.IsChecked = newFal;
            FciSorFbox.IsChecked = newFal;
            PrLoanForcebox.IsChecked = newFal;
            PrLoanLetterbox.IsChecked = newFal;
            PrLoanExpbox.IsChecked = newFal;
            FldFlagbox.IsChecked = newFal;
            Fcispelledbox.IsChecked = newFal;
            FciFbox.IsChecked = newFal;
            FciFAsBox.IsChecked = newFal;
            Fld250box.IsChecked = newFal;
            FldWbox.IsChecked = newFal;
            sfhaWbox.IsChecked = newFal;
            invalidFCIbox.IsChecked = newFal;
            lpDueExpbox.IsChecked = newFal;
            reoDueExpbox.IsChecked = newFal;
            auditPIFbox.IsChecked = newFal;
            auditVOLbox.IsChecked = newFal;
            pif8box.IsChecked = newFal;
            rolledpt1box.IsChecked = newFal;
            condo7312box.IsChecked = newFal;
            pay6box.IsChecked = newFal;
            pay7box.IsChecked = newFal;
            covGorJbox.IsChecked = newFal;
            covW250box.IsChecked = newFal;
            covE250box.IsChecked = newFal;
            dueExplinebox.IsChecked = newFal;
            expiredREObox.IsChecked = newFal;
            illhazCovbox.IsChecked = newFal;



        }

        private void wkBox_CheckedChanged(object sender, RoutedEventArgs e)
        {

            wkCheck.IsChecked = null;
            if ((HeaderFileHaBox.IsChecked == true) && (ExpLineBox.IsChecked == true) && (FciSorFbox.IsChecked == true) && (PrLoanForcebox.IsChecked == true) && (PrLoanLetterbox.IsChecked == true) && (PrLoanExpbox.IsChecked == true) && (FldFlagbox.IsChecked == true) && (Fcispelledbox.IsChecked == true) && (FciFbox.IsChecked == true) && (FciFAsBox.IsChecked == true) && (Fld250box.IsChecked == true) && (FldWbox.IsChecked == true) && (sfhaWbox.IsChecked == true) && (invalidFCIbox.IsChecked == true) && (lpDueExpbox.IsChecked == true) && (reoDueExpbox.IsChecked == true) && (auditPIFbox.IsChecked == true) && (auditVOLbox.IsChecked == true) && (pif8box.IsChecked == true) && (rolledpt1box.IsChecked == true) && (condo7312box.IsChecked == true) && (pay6box.IsChecked == true) && (pay7box.IsChecked == true) && (covGorJbox.IsChecked == true) && (covW250box.IsChecked == true) && (covE250box.IsChecked == true) && (dueExplinebox.IsChecked == true) && (expiredREObox.IsChecked == true) && (illhazCovbox.IsChecked == true))          
                wkCheck.IsChecked = true;
            if ((HeaderFileHaBox.IsChecked == false) && (ExpLineBox.IsChecked == false) && (FciSorFbox.IsChecked == false) && (PrLoanForcebox.IsChecked == false) && (PrLoanLetterbox.IsChecked == false) && (PrLoanExpbox.IsChecked == false) && (FldFlagbox.IsChecked == false) && (Fcispelledbox.IsChecked == false) && (FciFbox.IsChecked == false) && (FciFAsBox.IsChecked == false) && (Fld250box.IsChecked == false) && (FldWbox.IsChecked == false) && (sfhaWbox.IsChecked == false) && (invalidFCIbox.IsChecked == false) && (lpDueExpbox.IsChecked == false) && (reoDueExpbox.IsChecked == false) && (auditPIFbox.IsChecked == false) && (auditVOLbox.IsChecked == false) && (pif8box.IsChecked == false) && (rolledpt1box.IsChecked == false) && (condo7312box.IsChecked == false) && (pay6box.IsChecked == false) && (pay7box.IsChecked == false) && (covGorJbox.IsChecked == false) && (covW250box.IsChecked == false) && (covE250box.IsChecked == false) && (dueExplinebox.IsChecked == false) && (expiredREObox.IsChecked == false) && (illhazCovbox.IsChecked == false))
                wkCheck.IsChecked = false;
        }

        // Monthly reports checkbox to check and uncheck

        private void mnCheck_CheckedChanged(object sender, RoutedEventArgs e)
        {
            bool newMal = (mnCheck.IsChecked == true);
            AstoriaComExpUnder1box.IsChecked = newMal;
            ChoffBox.IsChecked = newMal;
            hazPolbox.IsChecked = newMal;
            revpremBox.IsChecked = newMal;
            revpremASbox.IsChecked = newMal;
            rev8box.IsChecked = newMal;
            secReviewbox.IsChecked = newMal;
            revPay2box.IsChecked = newMal;
            loan60daybox.IsChecked = newMal;
            
        }

        private void mnBox_CheckedChanged(object sender, RoutedEventArgs e)
        {

            mnCheck.IsChecked = null;
            if ((AstoriaComExpUnder1box.IsChecked == true) && (ChoffBox.IsChecked == true) && (hazPolbox.IsChecked == true) && (revpremBox.IsChecked == true) && (revpremASbox.IsChecked == true) && (rev8box.IsChecked == true) && (secReviewbox.IsChecked == true) && (revPay2box.IsChecked == true) && (loan60daybox.IsChecked == true))
                mnCheck.IsChecked = true;
            if ((AstoriaComExpUnder1box.IsChecked == false) && (ChoffBox.IsChecked == false) && (hazPolbox.IsChecked == false) && (revpremBox.IsChecked == false) && (revpremASbox.IsChecked == false) && (rev8box.IsChecked == false) && (secReviewbox.IsChecked == false) && (revPay2box.IsChecked == false) && (loan60daybox.IsChecked == false))
                    wkCheck.IsChecked = false;
        }



        // Loading report function

        public void btnLoad_Click(object sender, EventArgs e)
        {
            

          
            Excel.Range dateAstoriaComUnder1rng;
            DateTime? myDate = null;
            myDate = rptDate.SelectedDate;
            string rptYr = null;
            string mnFolder = null;
            string dayFolder = null;
            System.IO.FileInfo inforawAstoriaComUnder1 = new System.IO.FileInfo(@"N:\IRVS1012\IRVINE CLIENT SUPPORT\Clients\Dovenmuehle\z-Report Downloads\ASTORIA COMMERCIAL EXP UNDER 1 MILLION REPORT.xlsx");
            var infoloadAstoriaComUnder1 = System.IO.Path.Combine("dayDi", "ASTORIA COM EXP UNDER 1 MILLION REPORT.xlsx");
            DateTime? ctimeACU1 = inforawAstoriaComUnder1.CreationTime;
            rptYr = Convert.ToString(myDate);
            rptYr = myDate.Value.ToString("YYYY");
            mnFolder = Convert.ToString(myDate);
            mnFolder = myDate.Value.ToString("MMMM");
            dayFolder = Convert.ToString(myDate);
            dayFolder = myDate.Value.ToString("DD.MM.YY");
            var rptyrDi = System.IO.Path.Combine("mainRptDir", "+ rptYr+");
            var mnDi = System.IO.Path.Combine("rptyrDI", "+mnFolder+");
            var dayDi = System.IO.Path.Combine("mnDi", "+dayFolder+");
            



            if (readyCheck.IsChecked == true)
            {
                if (rptDate.SelectedDate == null)
                {
                    MessageBox.Show("Please select run date");
                    return;
                }

                
                // 242 report
                if (AstoriaComExpUnder1box.IsChecked == false)
                {
                    return;
                }
                

                if (ctimeACU1 != rptDate.SelectedDate)
                {
                    return;
                }


                if (!System.IO.Directory.Exists(rptyrDi))
                {
                    System.IO.Directory.CreateDirectory(rptyrDi);
                }

                if (!System.IO.Directory.Exists(mnDi))
                {
                    System.IO.Directory.CreateDirectory(mnDi);
                }

                if (!System.IO.Directory.Exists(dayDi))
                {
                    System.IO.Directory.CreateDirectory(dayDi);

                }

                if (!System.IO.Directory.Exists(infoloadAstoriaComUnder1))
                {
                    Excel.Application oXL = new Excel.Application();
                    oXL.Visible = true;
                    string wbrawAstoriaComUnder1 = @"N:\IRVS1012\IRVINE CLIENT SUPPORT\Clients\Dovenmuehle\z-Report Downloads\ASTORIA COMMERCIAL EXP UNDER 1 MILLION REPORT.xlsx";
                    Excel.Workbook wbAstoriaComUnder1 = oXL.Workbooks.Open(wbrawAstoriaComUnder1);
                    Excel.Worksheet sAstoriaComUnder1 = wbAstoriaComUnder1.Worksheets.get_Item(1);

                    sAstoriaComUnder1.Cells[1, 16] = "OPID";
                    sAstoriaComUnder1.Cells[1, 17] = "DATE COMPLETED";
                    sAstoriaComUnder1.Cells[1, 18] = "COMMENTS";

                    dateAstoriaComUnder1rng = sAstoriaComUnder1.get_Range("Q2", "Q1000");
                    dateAstoriaComUnder1rng.NumberFormat = "mm/dd/yyyy";

                    sAstoriaComUnder1.get_Range("A1", "R1").Font.Bold = true;
                    sAstoriaComUnder1.get_Range("A1", "R1").Interior.Color = Excel.XlRgbColor.rgbLightBlue;
                    sAstoriaComUnder1.get_Range("A1", "R1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    sAstoriaComUnder1.Columns.AutoFit();
                    wbAstoriaComUnder1.SaveAs(@"N:\IRVS1012\IRVINE CLIENT SUPPORT\Clients\Dovenmuehle\Daily ReportsTEST\Reports + rptYr +\+ mnFolder +\+ dayFolder +\ASTORIA COM EXP UNDER 1 MILLION REPORT.xlsx",
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    false,
                    false,
                    Excel.XlSaveAsAccessMode.xlShared,
                    false,
                    false,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value);



                    wbAstoriaComUnder1.Close();
                    oXL.Quit();
                    MessageBox.Show("Reports are loaded");
                }
            }




















































            //oWB = oXL.Workbooks.Add();
            //oSheet = (Excel._Worksheet)oWB.ActiveSheet;


            //string dXLPath = @"N:\IRVS1012\IRVINE CLIENT SUPPORT\Clients\Dovenmuehle\Raw Files\DAILY2 NONASTORIA BACKUP OPEN TASK BY HEADER.xlsx";
            //string oXLPath = @"N:\IRVS1012\IRVINE CLIENT SUPPORT\Clients\Dovenmuehle\Ready Files";

            //Excel.Workbook dWB = oXL.Workbooks.Open(dXLPath);
            //Excel.Worksheet dSheet = dWB.Worksheets.get_Item(1);


            //oSheet.Cells[1, 1] = "TASK ID";
            // oSheet.Cells[1, 2] = "TASK EXPECTED CLOSE DATE";
            //oSheet.Cells[1, 3] = "TASK RECEIVED DATE";
            // oSheet.Cells[1, 4] = "LOAN NUMBER";
            // oSheet.Cells[1, 5] = "TASK FOLLOW UP DATE";
            //  oSheet.Cells[1, 6] = "TASK RESPONSIBLE ID";
            //  oSheet.Cells[1, 7] = "TASK CONTACT USER ID";
            //  oSheet.Cells[1, 8] = "INVESTOR ID";
            //  oSheet.Cells[1, 9] = "OPID";
            //  oSheet.Cells[1, 10] = "DATE COMPLETED";
            //   oSheet.Cells[1, 11] = "COMMENTS";
            //   oSheet.Cells[1, 12] = "SLA";
            //  oSheet.Cells[1, 13] = "DAYS OPEN";






            //    Excel.Range filterRange = dSheet.get_Range("A1", "H10000");
            //   filterRange.AutoFilter(1, "HAZPOI", Excel.XlAutoFilterOperator.xlAnd);
            //    Excel.Range splRange = filterRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
            //    splRange.Copy();
            //    Excel.Range r1 = (Excel.Range)oSheet.Cells[1, 1];
            //    r1.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);



            //    bRng = oSheet.get_Range("B2", "B10000");
            //    bRng.NumberFormat = "mm/dd/yyyy";
            //     cRng = oSheet.get_Range("C2", "C10000");
            //     cRng.NumberFormat = "mm/dd/yyyy";
            //    eRng = oSheet.get_Range("E2", "E10000");
            //    eRng.NumberFormat = "mm/dd/yyyy";
            //    oRng = oSheet.get_Range("L2", "L1000");
            //    oRng.NumberFormat = "General";
            //    oRng.Formula = "=B2-C2";
            //     iRng = oSheet.get_Range("M2", "M1000");
            //    iRng.NumberFormat = "General";
            //     iRng.Formula = "=TODAY()-C2";

            //    oSheet.get_Range("A1", "M1").Font.Bold = true;
            //    oSheet.get_Range("A1", "M1").Interior.Color = Excel.XlRgbColor.rgbLightBlue;
            //    oSheet.get_Range("A1", "M1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            //     oSheet.Columns.AutoFit();

            //     dWB.Close(false);

            //    oWB.SaveAs(@"N:\IRVS1012\IRVINE CLIENT SUPPORT\Clients\Dovenmuehle\Ready Files\HAZPOI_TASKS.xlsx",
            //  System.Reflection.Missing.Value,
            //   System.Reflection.Missing.Value,
            //   System.Reflection.Missing.Value,
            //    false,
            //    false,
            //    Excel.XlSaveAsAccessMode.xlShared,
            //    false,
            //    false,
            //    System.Reflection.Missing.Value,
            //   System.Reflection.Missing.Value,
            //   System.Reflection.Missing.Value);
            //     oWB.Close();









            else if (readyCheck.IsChecked == false)
            {
                MessageBox.Show("Please select the checkbox when you're ready");
            }
            }












                }
    }



    

