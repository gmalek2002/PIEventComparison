using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OSIsoft.AF.PI;
using OSIsoft.AF.Data;
using OSIsoft.AF.Asset;
using OSIsoft.AF.Time;
using cxml = ClosedXML.Excel;

namespace SheetBreakAnalysis
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public const string tagsheetname = "Tags";
        public const string timesheetname = "Times";
        public const int topx = 2;                                         // If a variable/tag has a COV in the top "x" for that break, it is deemed significant.
        public static string topxheader = "In_Top_" + topx.ToString();

        // These variables will be used in the PIPointList.Summaries call, which is a part of the GetIntervals method in this module.
        public static AFSummaryTypes summarytypes = (AFSummaryTypes)174;
        public static AFCalculationBasis calc = AFCalculationBasis.TimeWeighted;
        public static AFTimestampCalculation timecalc = AFTimestampCalculation.Auto;
        public static PIPagingConfiguration pagingconfig = new PIPagingConfiguration(PIPageType.TagCount, 100);
        public static int numintervals = 1;



        // These are all the variables we want to capture in the summaries.  To be used as DataTable column names and spreadsheet column headers.
        public static string[] baseheader = new string[] { "Tag", "Start_Time", "End_Time" };
        public static string[] flags = summarytypes.ToString().Split(',');
        public static string[] extracalcs = new string[] { "COV", topxheader };
        public static string[] allsummaryheaders = baseheader.Union(flags).Union(extracalcs).ToArray();

        // These are the column headers for the top varying parameters report, and their data types.  To be used in CompileTopVariationReport method.
        public static string counttopname = "Count_Of_Top_" + topx.ToString();
        public static KeyValuePair<string, string> topvary1 = new KeyValuePair<string, string>("Tag", "System.String");
        public static KeyValuePair<string, string> topvary2 = new KeyValuePair<string, string>(counttopname, "System.Int32");
        public static Dictionary<string, string> topvaryingheaders = new Dictionary<string, string>() { {topvary1.Key,topvary1.Value },{topvary2.Key,topvary2.Value } };

        // Simple spreadsheet formatting preferences
        public static cxml.XLBorderStyleValues desiredborderstyle = cxml.XLBorderStyleValues.Thin;

        [STAThread]

        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        public static IList<string> GetTagNamesFromFile(cxml.IXLWorksheet sheet)
        {
            cxml.IXLCells tagcells = sheet.Column(1).CellsUsed();
            return tagcells.Select<cxml.IXLCell, string>(x => x.GetString()).ToList();
        }

        public static IList<DateTime> GetBreakTimesFromFile(cxml.IXLWorksheet sheet)
        {
            // Create a list of DateTimes based on the Excel timestamps in column A.
            // cxml refers to the ClosedXML package

            cxml.IXLCells tagcells = sheet.Column(1).CellsUsed();
            List<DateTime> breaktimes = tagcells.Select<cxml.IXLCell, DateTime>(x => x.GetDateTime()).ToList();
            breaktimes.Sort((x,y)=> x.CompareTo(y));
            return breaktimes;
        }

        public static cxml.XLWorkbook CreateOutputFile(bool includedebugsheet = false, string[] colheaders = null)
        {
            //If no column headers argument is passed, use the global headers array defined at top (which will be the case).
            if (colheaders == null || colheaders.Length == 0) { colheaders = allsummaryheaders; };

            cxml.XLWorkbook xlwbk = new cxml.XLWorkbook();
            cxml.IXLWorksheet firstsheet = xlwbk.Worksheets.Add("Summary");
            cxml.IXLWorksheet secondsheet = xlwbk.Worksheets.Add("Top Varying Parameters");
            
            if (includedebugsheet)
            {
                // The boolean argument gives option to add an extra sheet just for debugging.
                xlwbk.Worksheets.Add("Debug");
            }

            foreach(string flag in colheaders)
            {
                if(flag.First() == ' ')
                {
                    int flagindex = Array.IndexOf(colheaders, flag);
                    colheaders[flagindex] = flag.Substring(1);   // Remove heading blanks.
                }
            }

            firstsheet.Cell(1, 1).InsertData(colheaders, true);

            return xlwbk;

        }

        public static IList<(DateTime,DateTime)> CreateTimeRangesFromBreakTimes(IEnumerable<DateTime> breaktimes,int startsecs, int endsecs)
        {
            foreach(DateTime breaktime in breaktimes)
            {
                DateTime.SpecifyKind(breaktime, DateTimeKind.Local);
            }

            List<(DateTime, DateTime)> breaktimeranges = breaktimes.Select<DateTime, (DateTime,DateTime)>( x => (( x.AddSeconds(-1*startsecs), x.AddSeconds(-1 * endsecs) )) ).ToList();
            
            return breaktimeranges;

        }

        public static IList<IDictionary<AFSummaryTypes, AFValues>> GetSummaries(PIServer piserv, IEnumerable<string> tagnames, IEnumerable<(DateTime,DateTime)> timeranges)
        {

            List<AFTimeRange> aftimeranges = timeranges.Select<(DateTime, DateTime), AFTimeRange>(x => new AFTimeRange(x.Item1.ToString(),x.Item2.ToString())).ToList();
            aftimeranges.Sort((x,y)=>x.StartTime.CompareTo(y.StartTime));
            IList<AFTimeIntervalDefinition> intervals = Program.GetIntervals(aftimeranges,numintervals, false);

            PIPointList pointlist = new PIPointList(PIPoint.FindPIPoints(piserv, tagnames));
            IList<IDictionary<AFSummaryTypes, AFValues>> summaries = pointlist.Summaries(intervals,false,summarytypes,calc,timecalc,pagingconfig).ToList();

            return summaries;

        }

        public static IList<IDictionary<AFSummaryTypes, AFValues>> GetSummariesOneAtATime(PIServer piserv, IEnumerable<string> tagnames, IEnumerable<(DateTime, DateTime)> timeranges)
        {
            // This method is a slightly altered copy of the GetSummaries call defined above.  It performs a separate PIPointList.Summaries call for each time range,
            // rather than issuing one call for all time ranges at once.  This was done because the one call would calculate summary stats for all time ranges combined,
            // rather than calculate them for each time range separately.

            List<AFTimeRange> aftimeranges = timeranges.Select<(DateTime, DateTime), AFTimeRange>(x => new AFTimeRange(x.Item1.ToString(), x.Item2.ToString())).ToList();
            aftimeranges.Sort((x, y) => x.StartTime.CompareTo(y.StartTime));

            List<IDictionary<AFSummaryTypes, AFValues>> allsummaries = new List<IDictionary<AFSummaryTypes,AFValues>>();
            PIPointList pointlist = new PIPointList(PIPoint.FindPIPoints(piserv, tagnames));

            foreach (AFTimeRange aftimerng in aftimeranges)
            {
                IList<AFTimeIntervalDefinition> intervals = Program.GetIntervals(new List<AFTimeRange>() { aftimerng }, numintervals, false);
                IList<IDictionary<AFSummaryTypes, AFValues>> summaries = pointlist.Summaries(intervals, false, summarytypes, calc, timecalc, pagingconfig).ToList();
                allsummaries.AddRange(summaries);   
            }

            return allsummaries;

        }

        public static IList<IDictionary<AFSummaryTypes, AFValues>> GetSummaries(PIServer piserv, IEnumerable<string> tagnames, IEnumerable<(string, string)> timeranges)
        {

            List<AFTimeRange> aftimeranges = timeranges.Select<(string, string), AFTimeRange>(x => new AFTimeRange(x.Item1, x.Item2)).ToList();
            aftimeranges.Sort((x, y) => x.StartTime.CompareTo(y.StartTime));
            IList<AFTimeIntervalDefinition> intervals = Program.GetIntervals(aftimeranges, numintervals, false);

            PIPointList pointlist = new PIPointList(PIPoint.FindPIPoints(piserv, tagnames));
            IList<IDictionary<AFSummaryTypes, AFValues>> summaries = pointlist.Summaries(intervals, false, summarytypes, calc, timecalc, pagingconfig).ToList();

            return summaries;

        }

        public static List<AFTimeIntervalDefinition> GetIntervals(IEnumerable<AFTimeRange> aftimeranges, int numintervals, bool useselect = true)
        {

            if (useselect)
            {
                List<AFTimeIntervalDefinition> intervals = aftimeranges.Select<AFTimeRange, AFTimeIntervalDefinition>(x => new AFTimeIntervalDefinition(x, numintervals)).ToList();
                return intervals;
            }
            else
            {
                List<AFTimeIntervalDefinition> intervals = new List<AFTimeIntervalDefinition>();
                
                foreach(AFTimeRange aftimerange in aftimeranges)
                {
                    IList<AFTimeIntervalDefinition> theseintervals = new AFTimeSpan(aftimerange.Span).GetEvenTimeIntervalDefinitions(aftimerange);
                    intervals.AddRange(theseintervals);
                    
                }

                return intervals;
            }

        }



        public static DataTable FormatSummariesInDataTable(IEnumerable<IDictionary<AFSummaryTypes, AFValues>> summaries, cxml.IXLWorksheet xlsht, int intervalsecs)
        {
            // This method seems superior for pasting summary data into Excel so far.
            // The method "FormatSummariesForExcel" will likely be abandoned, as of this comment.

            DataTable dt = new DataTable();

            foreach(cxml.IXLCell headercell in xlsht.FirstRow().CellsUsed())
            {
                dt.Columns.Add(headercell.Value.ToString());
                if (headercell.Address.ColumnNumber > 3 & headercell.Value.ToString() != topxheader)
                {
                    dt.Columns[dt.Columns.Count-1].DataType = Type.GetType("System.Decimal");
                }
            }


            foreach (IDictionary<AFSummaryTypes,AFValues> summary in summaries)
            {
                DataRow newrow = dt.NewRow();

                AFTime starttime = summary.First().Value.First().Timestamp;
                AFTime endtime = starttime + new AFTimeSpan(0, 0, 0, 0, 0, intervalsecs, 0);
                newrow["Tag"] = summary.First().Value.PIPoint.Name;
                newrow["Start_Time"] = starttime.ToString();
                newrow["End_Time"] = endtime.ToString().Substring(0,endtime.ToString().IndexOf("M")+1);

                foreach (AFSummaryTypes summarytype in summary.Keys)
                {
                    newrow[summarytype.ToString()] = Convert.ToDouble(summary[summarytype].First().Value);
                }

                // The COV column must be calculated separately, because it is not returned with the PIPointList.Summaries call.
                // The "Top x?" column is then calculated based on COV and start time in a later method.
                newrow["COV"] = Convert.ToDouble(newrow["StdDev"]) / Convert.ToDouble(newrow["Average"]);

                dt.Rows.Add(newrow);
            }

            // Below is an attempt (which may or may not be retired) to add a top 10 column using DataTable instead of spreadsheet.
            // The spreadsheet option uses an identically named method, but is called within the FormatTableAfterPaste method.
            dt = AddTopTenColumn(dt);

            return dt;

        }

        public static void PasteSummariesIntoExcel(cxml.IXLWorksheet xlsht, DataTable summaries)
        {

            int currentrow = xlsht.LastRowUsed().RowNumber() + 1;
            xlsht.Cell(currentrow, 1).InsertData(summaries);

            // Add column for top 10 Yes/No.  There are two methods for this - one using spreadsheet argument, and the other taking a DataTable argument only.
            //AddTopTenColumn(xlsht, summaries, topx);
        }

        public static void FormatTableAfterPaste(cxml.IXLWorksheet xlsht, int numtags, int numcols)
        {
            // Below highlights summary rows for each break period a different color (light green or blue) to create visual separation between different break summaries.
            
            int numbreaktimes = (xlsht.RowsUsed().Count()-1)/numtags;
            bool colorswitch = false;

            foreach(int breaknum in Enumerable.Range(0,numbreaktimes))
            {
                if (colorswitch)
                {
                    xlsht.Range(xlsht.Cell(2 + breaknum * numtags, 1), xlsht.Cell(2 + (breaknum + 1) * numtags - 1, numcols)).Style.Fill.BackgroundColor = cxml.XLColor.LightBlue;
                }
                else
                {
                    xlsht.Range(xlsht.Cell(2 + breaknum * numtags, 1), xlsht.Cell(2 + (breaknum + 1) * numtags - 1, numcols)).Style.Fill.BackgroundColor = cxml.XLColor.LightGreen;
                }

                colorswitch = !colorswitch;
            }

            // Re-add borders and make header row bold.  Also adjust column width automatically.

            xlsht.RangeUsed().Style.Border.TopBorder = desiredborderstyle;
            xlsht.RangeUsed().Style.Border.LeftBorder = desiredborderstyle;
            xlsht.RangeUsed().Style.Border.RightBorder = desiredborderstyle;
            xlsht.RangeUsed().Style.Border.InsideBorder = desiredborderstyle;
            xlsht.RangeUsed().Style.Border.OutsideBorder = desiredborderstyle;

            xlsht.ColumnsUsed().AdjustToContents();
            xlsht.FirstRowUsed().Style.Font.Bold = true;

            // All columns except COV will be rounded to two decimal places.
            xlsht.RangeUsed().Columns(1,xlsht.ColumnsUsed().Count()-2).Style.NumberFormat.SetFormat("0.00");
            xlsht.Column(xlsht.ColumnsUsed().Count()-1).Style.NumberFormat.SetFormat("0.0000");

            // Sort results by start time, then by COV.
            xlsht.Range(2, 1, xlsht.RowsUsed().Count(), xlsht.ColumnsUsed().Count()).Sort("B,I",cxml.XLSortOrder.Descending);


        }

        public static DataTable AddTopTenColumn(DataTable dt)
        {

            int numrows = dt.Rows.Count;
            int covrank;

            foreach (DataRow dr in dt.Rows)
            {
                string dateval = dr["Start_Time"].ToString();
                string cov = dr["COV"].ToString();
                string filterstr = "COV > " + cov + " and Start_Time = '" + dateval + "'";
                covrank = dt.Select(filterstr).Count();
                if (covrank >= topx)
                {
                    //if the COV value for this row is greater than "x" in descending rank, then it is not top ten (or top "x").
                    dr[topxheader] = "No";
                }
                else
                {
                    dr[topxheader] = "Yes";
                }
            }

            return dt;
        }

        public static DataTable CompileTopVariationReport(IList<string> tagnames, DataTable summarytable)
        {
            DataTable dt = new DataTable("Top Variation Report");

            foreach(KeyValuePair<string,string> header in topvaryingheaders)
            {
                // Add columns using static headers defined globally.
                dt.Columns.Add(header.Key);
                dt.Columns[header.Key].DataType = Type.GetType(header.Value);
            }

            foreach(string tagname in tagnames)
            {
                DataRow dr = dt.NewRow();
                string filterstr = "Tag = '" + tagname + "' and " + topxheader + " = 'Yes'";  
                int topcount = summarytable.Select(filterstr).Count();                        // Count rows in summarytable where this tag was listed as a "top x" COV.

                dr["Tag"] = tagname;
                dr[counttopname] = topcount;

                dt.Rows.Add(dr);

            }

            return dt;
        }

        public static void PasteTopVariationReport(cxml.IXLWorksheet pastesht, DataTable dt) 
        {
            // Create top column headers using static global variable topvaryingheaders.  Change color, repair borders, make bold, etc
            pastesht.Cell(1, 1).InsertData(topvaryingheaders.Keys.ToArray(), true);
            cxml.IXLRange headerrng = pastesht.FirstRowUsed().Intersection(pastesht.RangeUsed()).AsRange();
            headerrng.Style.Fill.BackgroundColor = cxml.XLColor.LightCornflowerBlue;
            headerrng.Style.Border.TopBorder = desiredborderstyle;
            headerrng.Style.Border.BottomBorder = desiredborderstyle;
            headerrng.Style.Border.LeftBorder = desiredborderstyle;
            headerrng.Style.Border.RightBorder = desiredborderstyle;
            headerrng.Style.Border.OutsideBorder = desiredborderstyle;
            headerrng.Style.Font.Bold = true;

            // Paste Top Variation DataTable into the spreadsheet in the 2nd row.
            pastesht.Cell(2, 1).InsertData(dt);

            // Sort this sheet by the second column.  Make header bold and adjust column width.
            pastesht.Range(2, 1, pastesht.RowsUsed().Count(), pastesht.ColumnsUsed().Count()).Sort("B", cxml.XLSortOrder.Descending);
            pastesht.ColumnsUsed().AdjustToContents();
        }


        // Methods listed beyond this point are no longer used and are considered merely scrap for possible future endeavors.


        public static void AddTopTenColumn(cxml.IXLWorksheet xlsht, DataTable dt, int x = 10)
        {
            // Method 1 of 2 to add top ten column.  Attempt is made by manually inserting values into new spreadsheet column.
            // Also, the default value of 10 can be replaced by "x".

            int topxcolnum = 0;                                // column number for new top ten column;
            int covrank = 0;                                     // Will store zero-based rank for sorting by top 10.

            // If the header row already contains "Top 10?", record that column numer in the toptencolnum variable.
            // Otherwise, create a new column with that header.

            if (xlsht.FirstRowUsed().LastCellUsed().Value.ToString() == topxheader)
            {
                topxcolnum = xlsht.LastColumnUsed().ColumnNumber();
            }
            else
            {
                topxcolnum = xlsht.ColumnsUsed().Count() + 1;
                xlsht.Cell(1, topxcolnum).Value = topxheader;
            }

            foreach (int rownum in Enumerable.Range(2, xlsht.RangeUsed().Rows().Count() - 1))
            {
                string dateval = xlsht.Cell(rownum, 2).Value.ToString();
                string cov = xlsht.Cell(rownum, 9).Value.ToString();
                string filterstr = "COV > " + cov + " and Start_Time = '" + dateval + "'";
                covrank = dt.Select(filterstr).Count();
                if (covrank >= x)
                {
                    //if the COV value for this row is greater than "x" in descending rank, then it is not top ten (or top "x").
                    xlsht.Cell(rownum, topxcolnum).Value = "No";
                }
                else
                {
                    xlsht.Cell(rownum, topxcolnum).Value = "Yes";
                }
            }

        }

        public static void PasteSummariesIntoExcel(cxml.IXLWorksheet xlsht, IList<IDictionary<string, string>> summaries)
        {

            int currentrow = xlsht.LastRowUsed().RowNumber() + 1;
            int sumstatcol = 4;

            foreach (IDictionary<string, string> sumstat in summaries)
            {
                xlsht.Cell(currentrow, sumstatcol).InsertData(sumstat.Values, true);
                currentrow = xlsht.LastRowUsed().RowNumber() + 1;
            }

        }

        public static IList<IDictionary<string, string>> FormatSummariesForExcel(IEnumerable<IDictionary<AFSummaryTypes, AFValues>> summaries)
        {

            IList<IDictionary<string, string>> convertedsummarylist = new List<IDictionary<string, string>>();

            foreach (IDictionary<AFSummaryTypes, AFValues> tagsummary in summaries)
            {
                //IDictionary<string, string> convertedsummary = new Dictionary<string, string>();
                IDictionary<string, string> convertedsummary = tagsummary.ToDictionary<KeyValuePair<AFSummaryTypes, AFValues>, string, string>(x => x.Key.ToString(), x => x.Value.First().ToString());
                convertedsummarylist.Add(convertedsummary);
            }

            return convertedsummarylist;

        }

        public static IList<string> GetBreakTimeStringsFromFile(cxml.IXLWorksheet sheet)
        {
            cxml.IXLCells tagcells = sheet.Column(1).CellsUsed();
            return tagcells.Select<cxml.IXLCell, string>(x => x.GetString()).ToList();
        }

    }
}
