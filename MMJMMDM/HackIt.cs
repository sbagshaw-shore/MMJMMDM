using System;
using LinqToExcel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MMJMMDM
{
    public class HackIt
    {
        public IEnumerable<string> GetStringsFromFiles(IEnumerable<FileInfo> excelFiles)
        {
            var strings = new List<string>();

            foreach (var excelFile in excelFiles)
            {
                strings.AddRange(GetStringsFromFile(excelFile));
            }

            return strings;
        }

        private IEnumerable<string> GetStringsFromFile(FileInfo excelFile)
        {
            var strings = new List<string>();
            var excel = new ExcelQueryFactory(excelFile.FullName) { ReadOnly = true };

            var dailySheetRows = from c in excel.Worksheet("Daily").ToList()
                                 select c;

            var sleepSheetRows = from c in excel.Worksheet("Sleep Scores")
                                 select c;

            var dailyRowsGroupedByParticipant = dailySheetRows.ToList().GroupBy(x => x[GetColNumber("B")].Cast<string>()).ToList();

            foreach (var participantRows in dailyRowsGroupedByParticipant)
            {
                var participantSleepSheetRows = sleepSheetRows.ToList().
                                                Where(x => x[GetColNumber("B")] == participantRows.Key).ToList();

                strings.Add(GetStringForParticipant(participantRows, participantSleepSheetRows));
            }

            return strings;
        }

        private string GetStringForParticipant(IGrouping<string, Row> dailyRows, IEnumerable<Row> sleepRows)
        {
            // We’d be interested in the data from the 2nd day onwards for 7 days
            var dailyAnalysisRows = GetAnalysisRows(dailyRows.ToList(), 1, 7).ToList();

            var parts = new List<string>();

            parts.Add(GetParticipantId(dailyRows.Key)); // Participant_ID
            parts.Add(dailyAnalysisRows.First()[GetColNumber("G")]); // (start) Date

            foreach (var row in dailyAnalysisRows)
            {
                parts.Add(GetWeekendBit(row[GetColNumber("H")]));
            }

            parts.AddRange(GetSequenceFromWeek("BG", dailyAnalysisRows)); // WearTime_Min
            parts.AddRange(GetSequenceFromWeek("J", dailyAnalysisRows)); // KCAL
            parts.AddRange(GetSequenceFromWeek("BA", dailyAnalysisRows)); // STEPS
            parts.AddRange(GetSequenceFromWeek("L", dailyAnalysisRows)); // METS
            parts.AddRange(GetSequenceFromWeek("AH", dailyAnalysisRows)); // Total MVPA
            parts.AddRange(GetSequenceFromWeek("AI", dailyAnalysisRows)); // % in MVPA
            parts.AddRange(GetSequenceFromWeek("AJ", dailyAnalysisRows)); // MVPA per hour

            // these are not in sequence 1-7, but have a series of fields in each day...
            foreach (var row in dailyAnalysisRows)
            {
                parts.Add(row[GetColNumber("Z")]); // SED_MIN
                parts.Add(row[GetColNumber("AD")]); // SED_PCEN
                parts.Add(row[GetColNumber("AA")]); // LIGHT_MIN
                parts.Add(row[GetColNumber("AE")]); // LIGHT_PCEN
                parts.Add(row[GetColNumber("AB")]); // MOD_MIN
                parts.Add(row[GetColNumber("AF")]); // MOD_PCEN
                parts.Add(row[GetColNumber("AC")]); // VIG_MIN
                parts.Add(row[GetColNumber("AG")]); // VIG_PCEN
            }

            var sleepAnalysisRows = GetAnalysisRows(sleepRows.ToList(), 0, 8).ToList();

            parts.AddRange(GetSequenceFromWeek("Q", sleepAnalysisRows)); // TST 
            parts.AddRange(GetSequenceFromWeek("O", sleepAnalysisRows)); // EFF 
            parts.AddRange(GetSequenceFromWeek("S", sleepAnalysisRows)); // AWAKE_NUM 
            parts.AddRange(GetSequenceFromWeek("R", sleepAnalysisRows)); // WASO_MIN 

            parts = parts.Select(x => x.Replace(" %", string.Empty).Replace('/', '.')).ToList();

            return string.Join(",", parts);
        }

        private string GetParticipantId(string participantRowsKey)
        {
            return participantRowsKey.Replace(".agd", String.Empty).Replace("60sec", string.Empty);
        }

        private string GetWeekendBit(Cell cell)
        {
            var day = cell.Value.ToString();
            return day == "Saturday" || day == "Sunday" ? "1" : "0";
        }

        private IEnumerable<string> GetSequenceFromWeek(string columnName, IEnumerable<Row> rows)
        {
            foreach (var row in rows)
            {
                var isBlankRow = row.Count == 0;
                yield return isBlankRow ? string.Empty : row[GetColNumber(columnName)];
            }
        }

        private IEnumerable<Row> GetAnalysisRows(IEnumerable<Row> rows, int startIndex, int count)
        {
            if (rows.Count() < count)
            {
                var rowsToReturn = rows.ToList();

                for (var i = count - rows.Count(); i > 0; i--)
                {
                    rowsToReturn.Add(new Row()); // add in empty rows till the desired amount is reached...
                }

                return rowsToReturn;
            }

            return rows.ToList().GetRange(startIndex, count);
        }

        private static int GetColNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
            {
                return -1;
            }

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum - 1;
        }
    }
}
