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
    
    static class Program {

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

        static Func<A, R> Memoize<A, R>(this Func<A, R> f)
        {
            var map = new Dictionary<A, R>();
            return a =>
            {
                R value;
                if (map.TryGetValue(a, out value))
                    return value;
                value = f(a);
                map.Add(a, value);
                return value;
            };
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
                                --order by ISCED
                                ", ISCED_LEVEL, year),
                                temis);


            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

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
                    //int column = usedRange.Find(isced).Column;
                    int column = getCol(isced);

                    workSheet.Cells[row, column] = count;
                    Console.WriteLine(row.ToString() + " : " + column.ToString());
                }
            }
        }

        // A3: Number of students by level of education, age and sex

        static void sheetA3(Excel.Application excelApp, SqlConnection temis, String year)
        {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 29;     //row offset
            const int UNDER_TWO = 17;           //row 
            const int TWENTYFIVE_TWENTYNINE = 41;          //row
            const int OVER_TWENTYNINE = 42;        //row
            const int ZERO = 16;        //row

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A3"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

            SqlCommand cmd = new SqlCommand(
              string.Format(@"select ISCED, AGE, gender, count(1) as count from (
                                select 
                                --sg.class,
                                --CL.class_name,
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
                                  END,
                                AGE,
                                gender 
                                from v_sgca sg
                                join (select student_id, DATEDIFF(hour, sg.dob, '{0}/1/1' )/ 8766 AS AGE from v_sgca sg where year = {0}) AGES on AGES.student_id = sg.student_id
                                where year = {0}
                              ) 
                              v group by ISCED, AGE, gender
                              ", year),
                              temis);


            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string isced = rdr.GetString(0);
                    int age = rdr.GetInt32(1);
                    string gender = rdr.GetString(2);
                    int count = rdr.GetInt32(3);
                    Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", isced, gender, age, count));

                    int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
                    int row;
                    if(age >= 2 && age <= 24)
                    {
                        row = ZERO + age + rowOffset;
                    }
                    else if (age < 2)
                    {
                        row = UNDER_TWO + rowOffset;
                    }
                    else if (age >= 25 && age <= 29)
                    {
                        row = TWENTYFIVE_TWENTYNINE + rowOffset;
                    }
                    else if (age > 29)
                    {
                        row = OVER_TWENTYNINE + rowOffset;
                    }
                    else
                    {
                        Console.WriteLine("Invalid Age: " + age);
                        continue;
                    }

                    int column = getCol(isced);

                    workSheet.Cells[row, column] = workSheet.get_Range(GetCellAddress(column, row)).Value2 + count;
                    //Console.WriteLine(row.ToString() + " : " + column.ToString());
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

            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

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
                            if (male != 0) { workSheet.Cells[AGE_UNKNOWN, getCol(grade)] = male; }
                            if (female != 0) { workSheet.Cells[AGE_UNKNOWN + FEMALE_OFFSET, getCol(grade)] = female; }
                            continue;
                        }
                        int age = Convert.ToInt16(strAge);
                        if (age < 4) {
                            if (male != 0) { workSheet.Cells[UNDER_FOUR, getCol(grade)] = male; }
                            if (female != 0) { workSheet.Cells[UNDER_FOUR + FEMALE_OFFSET, getCol(grade)] = female; }
                            continue;
                        }

                        if (age > 24) {
                            if (male != 0) { workSheet.Cells[OVER_TWENTYFOUR, getCol(grade)] = male; }
                            if (female != 0) { workSheet.Cells[OVER_TWENTYFOUR + FEMALE_OFFSET, getCol(grade)] = female; }
                            continue;
                        }

                        // 4 <= age <= 24
                        if (male != 0) { workSheet.Cells[age + ZERO, getCol(grade)] = male; }
                        if (female != 0) { workSheet.Cells[age + ZERO + FEMALE_OFFSET, getCol(grade)] = female; }
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

            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

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
                            if (male != 0) { workSheet.Cells[AGE_UNKNOWN, getCol(grade)] = male; }
                            if (female != 0) { workSheet.Cells[AGE_UNKNOWN + FEMALE_OFFSET, getCol(grade)] = female; }
                            continue;
                        }
                        int age = Convert.ToInt16(strAge);
                        if (age < 10)
                        {
                            if (male != 0) { workSheet.Cells[UNDER_TEN, getCol(grade)] = male; }
                            if (female != 0) { workSheet.Cells[UNDER_TEN + FEMALE_OFFSET, getCol(grade)] = female; }
                            continue;
                        }

                        if (age > 24)
                        {
                            if (male != 0) { workSheet.Cells[OVER_TWENTYFOUR, getCol(grade)] = male; }
                            if (female != 0) { workSheet.Cells[OVER_TWENTYFOUR + FEMALE_OFFSET, getCol(grade)] = female; }
                            continue;
                        }

                        // 4 <= age <= 24
                        if (male != 0) { workSheet.Cells[age + ZERO, getCol(grade)] = male; }
                        if (female != 0) { workSheet.Cells[age + ZERO + FEMALE_OFFSET, getCol(grade)] = female; }
                    }
                }
            }
        }

        static void sheetA7(Excel.Application excelApp, SqlConnection temis, String year)
        {
            //Constant references for columns and rows            
            const int   MALE_ROW = 17;      //row
            const int FEMALE_ROW = 18;      //row
            const int ZERO = 14;            //column
            const int SECONDARY_OFFSET = 9; //column offset
            const int UPPER_SECONDARY = 68; //column

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A7"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            SqlCommand cmd = new SqlCommand(
              string.Format(@"select  
                                LEVEL.ISCED,
                                sg.class,
                                gender,
                                count(1) 
                                from v_sgca sg
                                join (select distinct sg.student_id as id,
	                                'ISCED' = 
		                                CASE
			                                WHEN sg.class >=1 and sg.class <=6  THEN 'ISCED 1'
			                                WHEN sg.class >=7 and sg.class <=8  THEN 'ISCED 2'
			                                WHEN sg.class >=9 and sg.class <=13 THEN 'ISCED 3'
		                                END from v_sgca sg where year = {0}) 
	                                LEVEL on LEVEL.id = sg.student_id
                                where year = {0}
                                and status = 'R'
                                and ISNULL(ISCED, 'NULL') != 'NULL'
                                group by ISCED, sg.class, gender
                              ", year),
                              temis);


            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string isced = rdr.GetString(0);
                    Decimal _class = rdr.GetDecimal(1);
                    string gender = rdr.GetString(2);
                    int count = rdr.GetInt32(3);
                    //Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", isced, gender, _class, count));

                    int row = gender == "M" ? MALE_ROW : FEMALE_ROW;
                    int column = ZERO + (int)(_class * 3);
                    column = (isced == "ISCED 2") ? column + SECONDARY_OFFSET : column;
                    column = (isced == "ISCED 3") ? UPPER_SECONDARY : column;
                    workSheet.Cells[row, column] = workSheet.get_Range(GetCellAddress(column, row)).Value2 + count;
                }
            }
        }

        static void sheetA8(Excel.Application excelApp, SqlConnection temis, String year)
        {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 20;     //row offset
            const int UNDER_FOUR = 17;           //row 
            const int OVER_EIGHTEEN = 33;        //row
            const int ZERO = 14;        //row
            const int AGE_UNKNOWN = 34;
            const int PRIMARY_COL = 17;
            const int LOWER_SECONDARY_COL = 23;

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A8"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

            SqlCommand cmd = new SqlCommand(
              string.Format(@"select 
                                LEVEL.ISCED,
                                Ages.AGE, 
                                gender,
                                count(1) as count
                                from V_SGCA sg
                                join (select student_id, DATEDIFF(hour, sg.dob, '{0}/1/1' )/ 8766 AS AGE from v_sgca sg) AGES on AGES.student_id = sg.student_id
                                join (select distinct sg.student_id as id,
	                                'ISCED' = 
		                                CASE
			                                WHEN sg.class >=1 and sg.class <=6 THEN 'PRIMARY'
			                                WHEN sg.class >=7 and sg.class <=8 THEN 'LOWER SECONDARY'
		                                END from v_sgca sg where year = {0}) 
	                                LEVEL on LEVEL.id = sg.student_id
                                where year = {0}
                                and status = 'N' -- New Entrant
                                and ISNULL(ISCED, 'NULL') != 'NULL'
                                group by AGE, gender, LEVEL.ISCED
                              ", year),
                              temis);


            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string isced = rdr.GetString(0);
                    int age = rdr.GetInt32(1);
                    string gender = rdr.GetString(2);
                    int count = rdr.GetInt32(3);
                    Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", isced, gender, age, count));

                    int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
                    int row;
                    if (age >= 2 && age <= 24)
                    {
                        row = ZERO + age + rowOffset;
                    }
                    else if (age < 4)
                    {
                        row = UNDER_FOUR + rowOffset;
                    }
                    else if (age > 18)
                    {
                        row = OVER_EIGHTEEN + rowOffset;
                    }
                    else
                    {
                        row = AGE_UNKNOWN + rowOffset;
                    }

                    int column = (isced == "PRIMARY") ? PRIMARY_COL : LOWER_SECONDARY_COL;

                    workSheet.Cells[row, column] = workSheet.get_Range(GetCellAddress(column, row)).Value2 + count;
                    Console.WriteLine(age + " " + isced + " " + gender);
                    Console.WriteLine(row.ToString() + " : " + column.ToString());
                }
            }
        }

        static void sheetA10(Excel.Application excelApp, SqlConnection temis, String year)
        {

            //Constant references for columns and rows            
            const int FEMALE_OFFSET = 4;     //row offset
            const int PUBLIC = 17;
            const int PRIVATE = 18;

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A10"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

            SqlCommand cmd = new SqlCommand(
              string.Format(@"select 
                                LEVEL.ISCED, 
                                school_type = CASE WHEN school_type = 1 THEN 'PUBLIC' WHEN school_type = 2 THEN 'PRIVATE' END,
                                gender, 
                                count(1) as count
                                from TGCA
                                left outer join STAFF on TGCA.staff_id = STAFF.staff_id
                                left outer join SCHOOLS on TGCA.school_id = SCHOOLS.school_id
                                left outer join (select distinct class,
	                                'ISCED' = 
		                                CASE
			                                WHEN class < 1 THEN 'ISCED 02'
			                                WHEN class >=1 and class <=6 THEN 'ISCED 1'
			                                WHEN class >=8.1 and class <=8.2 THEN 'ISCED 25'
			                                WHEN class >=7 and class <=8 THEN 'ISCED 24'
			                                WHEN class >=10.1 and class <=10.2 THEN 'ISCED 35'
			                                WHEN class >=9 and class <=13 THEN 'ISCED 34'
		                                END from TGCA where year = 2014 ) 
	                                LEVEL on LEVEL.class = TGCA.class
                                where year = 2014
                                and gender in ('F', 'M')
                                group by ISCED, gender, school_type
                              ", year),
                              temis);


            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string isced = rdr.GetString(0);
                    string schoolType = rdr.GetString(1);
                    string gender = rdr.GetString(2);
                    int count = rdr.GetInt32(3);
                    Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", isced, gender, schoolType, count));

                    int rowOffset = gender == "M" ? 0 : FEMALE_OFFSET;
                    int row = schoolType == "PUBLIC" ? PUBLIC : PRIVATE + rowOffset;

                    List<string> columns = new List<string>();
                    
                    if (isced == "ISCED 24" || isced == "ISCED 34")
                    {
                        columns.Add("ISCED 24+34");
                        columns.Add(isced.Substring(0, 7));
                    }
                    else if (isced == "ISCED 25" || isced == "ISCED 35")
                    {
                        columns.Add("ISCED 25+35");
                        columns.Add(isced.Substring(0, 7));
                    }
                    else
                    {
                        columns.Add(isced);
                    }
                    foreach (string column in columns) {
                        workSheet.Cells[row, getCol(column)] = workSheet.get_Range(GetCellAddress(getCol(column), row)).Value2 + count;
                        //Console.WriteLine(age + " " + isced + " " + gender);
                        //Console.WriteLine(row.ToString() + " : " + column.ToString());
                    }
                }
            }
        }

        static void sheetA12(Excel.Application excelApp, SqlConnection temis, String year)
        {

            //Constant references for columns and rows            
            const int FEMALE = 18;     //row offset
            const int MALE = 17;

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets["A12"];
            workSheet.Activate();
            Excel.Range usedRange = workSheet.UsedRange;

            Func<string, int> getCol = null;
            getCol = n => usedRange.Find(n).Column;
            getCol.Memoize();

            SqlCommand cmd = new SqlCommand(
              string.Format(@"select 
                                LEVEL.ISCED, 
                                gender, 
                                count(1)
                                from TGCA
                                left outer join STAFF on TGCA.staff_id = STAFF.staff_id
                                left outer join (select distinct class,
	                                'ISCED' = 
		                                CASE
			                                WHEN class < 1 THEN 'ISCED 02'
			                                WHEN class >=1 and class <=6 THEN 'ISCED 1'
			                                WHEN class >=8.1 and class <=8.2 THEN 'ISCED 25'
			                                WHEN class >=7 and class <=8 THEN 'ISCED 24'
			                                WHEN class >=10.1 and class <=10.2 THEN 'ISCED 35'
			                                WHEN class >=9 and class <=13 THEN 'ISCED 34'
		                                END from TGCA where year = {0}) 
	                                LEVEL on LEVEL.class = TGCA.class
                                where year = {0}
                                and STAFF.teaching_qual = 'Y'
                                group by ISCED, gender
                              ", year),
                              temis);


            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    string isced = rdr.GetString(0);
                    string gender = rdr.GetString(1);
                    int count = rdr.GetInt32(2);
                    Console.WriteLine(String.Format("{0}, {1}, {2}", isced, gender, count));

                    int row = gender == "M" ? MALE : FEMALE;

                    List<string> columns = new List<string>();

                    if (isced == "ISCED 24" || isced == "ISCED 34")
                    {
                        columns.Add("ISCED 24+34");
                        columns.Add(isced.Substring(0, 7));
                    }
                    else if (isced == "ISCED 25" || isced == "ISCED 35")
                    {
                        columns.Add("ISCED 25+35");
                        columns.Add(isced.Substring(0, 7));
                    }
                    else
                    {
                        columns.Add(isced);
                    }
                    foreach (string column in columns)
                    {
                        workSheet.Cells[row, getCol(column)] = workSheet.get_Range(GetCellAddress(getCol(column), row)).Value2 + count;
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
            excelApp.Visible = true;
            sheetA2(excelApp, temis, "2014");
            sheetA3(excelApp, temis, "2014");
            sheetA5(excelApp, temis);
            sheetA6(excelApp, temis);
            sheetA7(excelApp, temis, "2014");
            sheetA8(excelApp, temis, "2014");
            sheetA10(excelApp, temis, "2014");
            sheetA12(excelApp, temis, "2014");

            excelApp.Visible = true;
            Console.ReadKey();
        }
    }
}
