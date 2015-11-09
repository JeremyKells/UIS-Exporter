using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace EmisExporter { 
    class Program {

        
        static string GetCellAddress(int col, int row) {
            StringBuilder sb = new StringBuilder();
            do {
                col--;
                sb.Insert(0, (char)('A' + (col % 26)));
                col /= 26;
            } while (col > 0);
            sb.Append(row);
            return sb.ToString();
        }

        // A5: Number of students in initial primary education by age, grade and sex																													
        static void sheetA5(Excel.Application excelApp, SqlConnection temis) {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 26;     //row offset
            const int AGE_UNKNOWN = 40;       //row
            const int UNDER_FOUR = 17;        //row
            const int OVER_TWENTYFOUR = 39;   //row
            const int ZERO = 14;              //row offset 
            const int UNSPECIFIED_GRADE = 38; //column AL

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A5"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            SqlCommand cmd = new SqlCommand("sp_rpt_agewise_enrolment", temis);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@TYPE", "N"));
            cmd.Parameters.Add(new SqlParameter("@Year", 2015));

            using (SqlDataReader rdr = cmd.ExecuteReader()) {
                while (rdr.Read()) {
                    string grade = (string)rdr["class_name"];
                    if (grade.IndexOf("Class ") != -1 && Convert.ToDouble(rdr["class"]) <= 6) {
                        string strAge = (string)rdr["AGE"];
                        int male = (int)rdr["male"];
                        int female = (int)rdr["female"];
                        grade = grade.Replace("Class", "Grade");
                        //decimal _class = (decimal)rdr["class"];
                        //String className = (String)rdr["class_name"];
                        //Console.WriteLine(
                        //    _class + ", " +
                        //    className + ", " +
                        //    strAge + ", " +
                        //    male + ", " +
                        //    female);

                        if (strAge == "N/A") {
                            if (male != 0) { workSheet.Cells[AGE_UNKNOWN, usedRange.Find(grade).Column] = male; }
                            if (female != 0) { workSheet.Cells[AGE_UNKNOWN + FEMALE_OFFSET, usedRange.Find(grade).Column] = female; }
                            continue;
                        }
                        int age = Convert.ToInt16(strAge);
                        if (age < 4) {
                            if (male != 0) { workSheet.Cells[UNDER_FOUR, usedRange.Find(grade).Column] = male; }
                            if (female != 0) { workSheet.Cells[UNDER_FOUR + FEMALE_OFFSET, usedRange.Find(grade).Column] = female; }
                            continue;
                        }

                        if (age > 24) {
                            if (male != 0) { workSheet.Cells[OVER_TWENTYFOUR, usedRange.Find(grade).Column] = male; }
                            if (female != 0) { workSheet.Cells[OVER_TWENTYFOUR + FEMALE_OFFSET, usedRange.Find(grade).Column] = female; }
                            continue;
                        }

                        // 4 <= age <= 24
                        if (male != 0) { workSheet.Cells[age + ZERO, usedRange.Find(grade).Column] = male; }
                        if (female != 0) { workSheet.Cells[age + ZERO + FEMALE_OFFSET, usedRange.Find(grade).Column] = female; }
                    }
                }
            }

            // Tidy Empty Values
            for (int row = UNDER_FOUR; row <= AGE_UNKNOWN; row++) {
                for (int grade = 0; grade <= 7; grade++) {

                    int column = (grade == 0) ? UNSPECIFIED_GRADE : usedRange.Find("Grade " + grade).Column;
                    int[] rows = { row, row + FEMALE_OFFSET };
                    foreach (int _row in rows) {
                        if (workSheet.get_Range(GetCellAddress(column, _row)).Value2 == null) {
                            workSheet.Cells[_row, column] = 0;
                            //workSheet.Cells[_row, column + 1] = "Z";
                        }
                    }
                }
            }
        }

        // A6: Number of students in initial lower secondary general education by age, grade and sex																										
        static void sheetA6(Excel.Application excelApp, SqlConnection temis)
        {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 20;     //row offset
            const int AGE_UNKNOWN = 34;       //row
            const int UNDER_TEN = 17;         //row
            const int OVER_TWENTYFOUR = 33;   //row
            const int ZERO = 8;               //row offset 
            const int UNSPECIFIED_GRADE = 35; // column AI

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A6"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            SqlCommand cmd = new SqlCommand("sp_rpt_agewise_enrolment", temis);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@TYPE", "N"));
            cmd.Parameters.Add(new SqlParameter("@Year", 2015));

            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string grade = (string)rdr["class_name"];
                    if (grade.IndexOf("Class ") != -1 && Convert.ToDouble(rdr["class"]) >= 7)
                    {
                        string strAge = (string)rdr["AGE"];
                        int male = (int)rdr["male"];
                        int female = (int)rdr["female"];
                        decimal _class = (decimal)rdr["class"];
                        grade = "Grade " + ((int)_class - 6).ToString();
                        //String className = (String)rdr["class_name"];
                        //Console.WriteLine(
                        //    _class + ", " +
                        //    className + ", " +
                        //    strAge + ", " +
                        //    male + ", " +
                        //    female);

                        if (strAge == "N/A")
                        {
                            if (male != 0) { workSheet.Cells[AGE_UNKNOWN, usedRange.Find(grade).Column] = male; }
                            if (female != 0) { workSheet.Cells[AGE_UNKNOWN + FEMALE_OFFSET, usedRange.Find(grade).Column] = female; }
                            continue;
                        }
                        int age = Convert.ToInt16(strAge);
                        if (age < 10)
                        {
                            if (male != 0) { workSheet.Cells[UNDER_TEN, usedRange.Find(grade).Column] = male; }
                            if (female != 0) { workSheet.Cells[UNDER_TEN + FEMALE_OFFSET, usedRange.Find(grade).Column] = female; }
                            continue;
                        }

                        if (age > 24)
                        {
                            if (male != 0) { workSheet.Cells[OVER_TWENTYFOUR, usedRange.Find(grade).Column] = male; }
                            if (female != 0) { workSheet.Cells[OVER_TWENTYFOUR + FEMALE_OFFSET, usedRange.Find(grade).Column] = female; }
                            continue;
                        }

                        // 4 <= age <= 24
                        if (male != 0) { workSheet.Cells[age + ZERO, usedRange.Find(grade).Column] = male; }
                        if (female != 0) { workSheet.Cells[age + ZERO + FEMALE_OFFSET, usedRange.Find(grade).Column] = female; }
                    }
                }
            }

            // Tidy Empty Values
            for (int row = UNDER_TEN; row <= AGE_UNKNOWN; row++)
            {
                for (int grade = 0; grade <= 6; grade++)
                {

                    int column = (grade == 0) ? UNSPECIFIED_GRADE : usedRange.Find("Grade " + grade).Column;
                    int[] rows = { row, row + FEMALE_OFFSET };
                    foreach (int _row in rows)
                    {
                        if (workSheet.get_Range(GetCellAddress(column, _row)).Value2 == null)
                        {
                            workSheet.Cells[_row, column] = 0;
                            //workSheet.Cells[_row, column + 1] = "Z";
                        }
                    }
                }
            }
        }

        static void Main(string[] args) {
            var excelApp = new Excel.Application();
            excelApp.Workbooks.Add("D:\\EMIS\\UIS_ED_A_2015.xlsx");

            SqlConnection temis = new SqlConnection();
            temis.ConnectionString =
                "Data Source=WIN-LV30Q52A0DN\\SQLEXPRESS;" +
                "Initial Catalog=TEMISS;" +
                "User id=sa;" +
                "Password=ab1234;";
            temis.Open();
            sheetA5(excelApp, temis);
            sheetA6(excelApp, temis);
            excelApp.Visible = true;
        }
    }
}
