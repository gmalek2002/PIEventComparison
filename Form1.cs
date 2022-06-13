using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using OSIsoft.AF.UI;
using OSIsoft.AF.Time;
using OSIsoft.AF.PI;
using OSIsoft.AF.Data;
using OSIsoft.AF.Asset;
using cxml = ClosedXML.Excel;

namespace SheetBreakAnalysis
{
    public partial class Form1 : Form
    {

        public string xlwbkname = "Sheet Break Analysis on " + DateTime.Now.ToString().Replace("/", "-").Replace(":", ",");
        //public string savepath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\" + xlwbkname + ".xlsm";

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //string endmessage = Testing.TestSummaries();
            //this.textBox6.Text = endmessage;
            DebugCopy.ConfirmedWorking(new string[] { "" });

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox1.Checked)
            {
                // If sheet break times are included in file, then start/end time, and sheet break tag/value boxes are no longer necessary.
                this.textBox1.Enabled = false;
                this.textBox2.Enabled = false;
                this.textBox3.Enabled = false;
                this.textBox4.Enabled = false;
            }
            else
            {
                this.textBox1.Enabled = true;
                this.textBox2.Enabled = true;
                this.textBox3.Enabled = true;
                this.textBox4.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        { 
            // Begin timer
            var watch = new Stopwatch();
            watch.Start();

            this.textBox6.Text = "Program has begun.";

            //Set some basic parameters.
            PIServer piserv = this.piServerPicker1.PIServer;
            string savepath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\" + xlwbkname + ".xlsm";

            cxml.XLWorkbook inputwbk = new cxml.XLWorkbook(this.textBox5.Text);
            IList<string> tagnames = Program.GetTagNamesFromFile(inputwbk.Worksheet(1));

            // Determine startsecs and endsecs.
            int startsecs = 630; //Default
            int endsecs = 30;    //Default
            try
            {
                // Convert user input strings to integers.
                startsecs = Int32.Parse(this.textBox8.Text);
                endsecs = Int32.Parse(this.textBox7.Text);
            }
            catch (FormatException fe)
            {
                this.textBox6.Text = "Program exited with error: start and/or end time formatted incorrectly.";
                return;
            }

            if(endsecs > startsecs)
            {
                //This would imply that the analysis start time is later than the end time, which can't be allowed.
                this.textBox6.Text = "Program exited with error: Analysis Window End should be less than start.";
                return;
            }

            if (!this.checkBox1.Checked)
            {
                // If user selects automatic break time option, tell them it's not enabled yet and exit (will remove this once feature is enabled).
                this.textBox6.Text = "Automated sheet break times currently not enabled.  Exiting...";
                return;
            }

            IList<(DateTime, DateTime)> timeranges = Program.CreateTimeRangesFromBreakTimes(Program.GetBreakTimesFromFile(inputwbk.Worksheet(2)), startsecs, endsecs);
            inputwbk.Save();

            //Now assemble output file, then acquire summary values, etc.

            cxml.XLWorkbook outputwbk = Program.CreateOutputFile();
            IList<IDictionary<AFSummaryTypes, AFValues>> summaries = Program.GetSummariesOneAtATime(piserv,tagnames,timeranges);

            // Create a DataTable using summaries dictionary.  Then paste into excel table and format the excel table.
            DataTable formattedsummaries = Program.FormatSummariesInDataTable(summaries, outputwbk.Worksheet(1),(startsecs - endsecs));
            Program.PasteSummariesIntoExcel(outputwbk.Worksheet("Summary"), formattedsummaries);
            Program.FormatTableAfterPaste(outputwbk.Worksheet("Summary"), tagnames.Count, formattedsummaries.Columns.Count);

            // Create a DataTable for the "Top Varying Parameters" report and insert into the second tab, then format, etc.
            DataTable topvarsummary = Program.CompileTopVariationReport(tagnames, formattedsummaries);
            Program.PasteTopVariationReport(outputwbk.Worksheet("Top Varying Parameters"), topvarsummary);
            outputwbk.SaveAs(savepath);

            watch.Stop();
            this.textBox6.Text = "Program finished in " +  System.Decimal.Ceiling(Convert.ToDecimal(watch.Elapsed.TotalSeconds))  + " seconds.  Check Documents folder.";

        }
    }
}

