-- A4 ?? No stats on Students in formal adult education

-- A9 No Data

-- A11 Trained Teachers???


-- A2
select  
LEVEL.ISCED,
sg.school_type,
gender,
count(1)
from v_sgca sg
join (select distinct sg.class as class,
	'ISCED' = 
		CASE
			WHEN sg.class < 1 THEN 'PRESCHOOL'
			WHEN sg.class >=1 and sg.class <=6 THEN 'PRIMARY'
			WHEN sg.class >=7 and sg.class <=8 THEN 'LOWER SECONDARY'
			WHEN sg.class >=9 and sg.class <=13 THEN 'UPPER SECONDARY'
		END from v_sgca sg ) 
	LEVEL on LEVEL.class = sg.class
where year = 2015
group by ISCED, sg.school_type, gender
order by ISCED

-- A3 A5 A6 
select 
ISCED = 
	CASE
		WHEN sg.class < 1 THEN 'PRESCHOOL'
		WHEN sg.class >=1 and sg.class <=6 THEN 'PRIMARY'
		WHEN sg.class >=7 and sg.class <=8 THEN 'LOWER SECONDARY'
		WHEN sg.class >=9 and sg.class <=13 THEN 'UPPER SECONDARY'
	END,
sg.class,
CL.class_name,
AGES.AGE,
sum(case when sg.gender = 'M' then 1 else 0 end) Male,
sum(case when sg.gender = 'F' then 1 else 0 end) Female
from v_sgca sg
join CLASSES_LEVELS CL on CL.class = sg.class
join (select student_id, DATEDIFF(hour, sg.dob, '2015/1/1' )/ 8766 AS AGE from v_sgca sg) AGES on AGES.student_id = sg.student_id
where year = 2015
group by sg.class, AGES.AGE, CL.class_name
order by sg.class, CL.class_name, AGES.AGE

select sg.class, 
sum(case when sg.gender = 'M' then 1 else 0 end) Male,
sum(case when sg.gender = 'F' then 1 else 0 end) Female
from V_SGCA sg
where status = 'R'
and year = 2015
group by sg.class


-- A7
select  
LEVEL.ISCED,
sum(case when sg.gender = 'M' then 1 else 0 end) Male,
sum(case when sg.gender = 'F' then 1 else 0 end) Female
from v_sgca sg
join (select distinct sg.student_id as id,
	'ISCED' = 
		CASE
			WHEN sg.class < 1 THEN 'PRESCHOOL'
			WHEN sg.class >=1 and sg.class <=6 THEN 'PRIMARY'
			WHEN sg.class >=7 and sg.class <=8 THEN 'LOWER SECONDARY'
			WHEN sg.class >=9 and sg.class <=13 THEN 'UPPER SECONDARY'
		END from v_sgca sg where year = 2015) 
	LEVEL on LEVEL.id = sg.student_id
where year = 2015
group by ISCED




-- A8, no info on who attenced ECE
select 
class, 
count(Ages.AGE)
from V_SGCA sg
join (select student_id, DATEDIFF(hour, sg.dob, '2015/1/1' )/ 8766 AS AGE from v_sgca sg) AGES on AGES.student_id = sg.student_id
where year = 2015
and status = 'N' -- New Entrant
group by class
order by class

-- A10
select 
LEVEL.ISCED, 
gender, 
school_type,
count(1)
from TGCA
left outer join STAFF on TGCA.staff_id = STAFF.staff_id
left outer join SCHOOLS on TGCA.school_id = SCHOOLS.school_id
left outer join (select distinct class,
	'ISCED' = 
		CASE
			WHEN class < 1 THEN 'PRESCHOOL'
			WHEN class >=1 and class <=6 THEN 'PRIMARY'
			WHEN class >=7 and class <=8 THEN 'LOWER SECONDARY'
			WHEN class >=9 and class <=13 THEN 'UPPER SECONDARY'
		END from TGCA ) 
	LEVEL on LEVEL.class = TGCA.class
where year = 2015
group by ISCED, gender, school_type


-- A12 (Qualified)  
select 
LEVEL.ISCED, 
gender, 
--STAFF.qual_level,
--STAFF.teaching_qual,
--STAFF.qual_name,
count(1)
from TGCA
left outer join STAFF on TGCA.staff_id = STAFF.staff_id
left outer join (select distinct class,
	'ISCED' = 
		CASE
			WHEN class < 1 THEN 'PRESCHOOL'
			WHEN class >=1 and class <=6 THEN 'PRIMARY'
			WHEN class >=7 and class <=8 THEN 'LOWER SECONDARY'
			WHEN class >=9 and class <=13 THEN 'UPPER SECONDARY'
		END from TGCA ) 
	LEVEL on LEVEL.class = TGCA.class
where year = 2015
and STAFF.teaching_qual = 'Y'
group by ISCED, gender
--,STAFF.qual_level,
--STAFF.teaching_qual,
--STAFF.qual_name

