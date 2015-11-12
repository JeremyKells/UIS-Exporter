using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;


//select sg.class,
//CL.class_name,
//--sg.student_id,
//AGES.AGE,
//sum(case when sg.gender = 'M' then 1 else 0 end) Male,
//sum(case when sg.gender = 'F' then 1 else 0 end) Female
//from v_sgca sg
//join CLASSES_LEVELS CL on CL.class = sg.class
//join (select student_id, DATEDIFF(hour, sg.dob, '2015/1/1' )/ 8766 AS AGE from v_sgca sg) AGES on AGES.student_id = sg.student_id
//where year = 2015
//group by sg.class, AGES.AGE, CL.class_name
//order by sg.class, CL.class_name, AGES.AGE


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


        const string ISCED_LEVEL = @"(  select distinct sg.class, class_name,
                                        'ISCED' = 
		                                        CASE
			                                        -- Currently K1, K2, K3 all ISCED 02     THEN 'ISCED 01'
			                                        WHEN sg.class < 1 THEN 'ISCED 02'
			                                        WHEN sg.class >=1 and sg.class <=6 THEN 'ISCED 1'
			                                        WHEN sg.class >=7 and sg.class <=8 THEN 'ISCED 24'
			                                        WHEN sg.class = 8.1 or sg.class = 8.2 THEN 'ISCED 25'
			                                        WHEN sg.class = 10.1 or sg.class = 10.2 THEN 'ISCED 35'
			                                        WHEN sg.class >=9 and sg.class <=13 THEN 'ISCED 34'
			                                        -- Currently no data for ISCED 44 / ISCED 45
		                                        END 
                                         from v_sgca sg )";

        // A2: Number of students by level of education, intensity of participation, type of institution and sex																																																			
        static void sheetA2(Excel.Application excelApp, SqlConnection temis, String year)
        {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 4;     //row offset
            const int PUBLIC = 17;           //row 
            const int PRIVATE = 18;          //row
            const int PART_TIME = 28;        //row

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A2"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            SqlCommand cmd = new SqlCommand(
              string.Format(@"  select
                                LEVEL.ISCED,
                                'schoolType' = case when sg.school_type = 1 then 'Public' else 'Private' end,
                                sg.gender,
                                count(1) as count
                                from v_sgca sg
                                join {0} LEVEL on LEVEL.class = sg.class
                                where year = {1}
                                group by ISCED, school_type, gender
                                order by ISCED
                                ", ISCED_LEVEL, year),
                                temis);

            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    if (rdr.IsDBNull(2))
                    {
                        Console.WriteLine("Skipping row, count: " + rdr.GetInt32(3).ToString());
                        continue;
                    }
                    string isced = rdr.GetString(0);
                    string schoolType = rdr.GetString(1);
                    string gender = rdr.GetString(2);
                    int count = rdr.GetInt32(3);
                    Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", isced, gender, schoolType, count));

                    int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
                    int row = (schoolType == "Public" ? PUBLIC : PRIVATE) + rowOffset;
                    int column = usedRange.Find(isced).Column;

                    workSheet.Cells[row, column] = count;
                    Console.WriteLine(row.ToString() + " : " + column.ToString());
                }
            }
            // Tidy empty cells
            String[] ISCEDs = new string[] {    "ISCED 01",
                                                "ISCED 02",
                                                "ISCED 1",
                                                "ISCED 24",
                                                "ISCED 25",
                                                "ISCED 34",
                                                "ISCED 35",
                                                "ISCED 44",
                                                "ISCED 45" };

            int[] genderOffsets = new int[] { 0, FEMALE_OFFSET };
            int[] schoolTypes = new int[] { PRIVATE, PUBLIC};
            foreach(string isced in ISCEDs)
            {
                int column = usedRange.Find(isced).Column;
                foreach (int gender in genderOffsets)
                {
                    workSheet.Cells[PART_TIME, column] = 0;   // No Part-time Teachers
                    foreach (int schoolType in schoolTypes)
                    {
                        int row = schoolType + gender;
                        
                        if (workSheet.get_Range(GetCellAddress(column, row)).Value2 == null)
                        {
                            workSheet.Cells[row, column] = 0;
                        }

                    }
                }

            }
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
            cmd.Parameters.Add(new SqlParameter("@Year", 2014));

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
            cmd.Parameters.Add(new SqlParameter("@Year", 2014));

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
            sheetA2(excelApp, temis, "2014");
            sheetA5(excelApp, temis);
            sheetA6(excelApp, temis);
            excelApp.Visible = true;
            Console.ReadKey();
        }
    }
}
