USE [POS_LOCAL]
GO

/****** Object:  StoredProcedure [dbo].[spp_DownLoadCpoint]    Script Date: 10/16/2012 12:57:10 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




 

alter PROCEDURE [dbo].[spp_DownLoadCpoint]
 

@ServerName char(20)
,@DBServer char(20) 

AS


declare @SQL nvarchar(2000) 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE card
INSERT card ( Card_Nr, Cust_Nr, Card_Expired_Date, Card_Activate_Date, Card_Point, Card_Status, Card_Claim_Date, User_ID_Claim, User_ID_Activate, 
             Data_Status, Update_Status, rowguid)
SELECT  Card_Nr, Cust_Nr, Card_Expired_Date, Card_Activate_Date, Card_Point, Card_Status, Card_Claim_Date, User_ID_Claim, User_ID_Activate, 
       Data_Status, Update_Status, rowguid
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.card'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Param_AddPoint
INSERT Customer_Param_AddPoint (Start_Date, End_Date)
SELECT Start_Date, End_Date
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Customer_Param_AddPoint'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Param_Day_MemberCard
INSERT Customer_Param_Day_MemberCard (Day_Code, Day_Flag)
SELECT Day_Code, Day_Flag
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Customer_Param_Day_MemberCard'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Param_Amount_MemberCard_Min_Purch
INSERT Customer_Param_Amount_MemberCard_Min_Purch ( Min)
SELECT  Min
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Customer_Param_Amount_MemberCard_Min_Purch'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Param_Amount_MemberCard
INSERT Customer_Param_Amount_MemberCard (Amount, Multiple, Lock)
SELECT  Amount, Multiple, Lock
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Customer_Param_Amount_MemberCard'
Execute sp_executesql @SQL 


Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Param_Voc_GetPoint
INSERT Customer_Param_Voc_GetPoint (Voc, Date_Start, Date_Finish)
SELECT Voc, Date_Start, Date_Finish
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Customer_Param_Voc_GetPoint'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Param_Bonus_PointCC
INSERT Customer_Param_Bonus_PointCC (Nomor, Date_Update, Date_Start, Date_Finish, Event, Point, CC, NoCC, Amount, Total_Amount, User_ID, Time_Update, Keterangan, 
									Time_Start, Time_Finish, Operator, Flag, Branch)
SELECT Nomor, Date_Update, Date_Start, Date_Finish, Event, Point, CC, NoCC, Amount, Total_Amount, User_ID, Time_Update, Keterangan, 
       Time_Start, Time_Finish, Operator, Flag, Branch
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Customer_Param_Bonus_PointCC'
Execute sp_executesql @SQL 


Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Param_AddPoint
INSERT Customer_Param_AddPoint(Start_Date, End_Date)
SELECT Start_Date, End_Date
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Customer_Param_AddPoint'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Param_Day
INSERT Customer_Param_Day ( Day_Code, Day_Flag)
SELECT  Day_Code, Day_Flag
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Customer_Param_Day'
Execute sp_executesql @SQL 



Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Param_Amount
INSERT Customer_Param_Amount (Amount, Multiple, Lock)
SELECT Amount, Multiple, Lock
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Customer_Param_Amount'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Cust_Param_Bonus
INSERT Cust_Param_Bonus (Jenis_Kartu, Event_Name, Point, Start, Finish, ActiveDay, Branch, Status_Active)
SELECT Jenis_Kartu, Event_Name, Point, Start, Finish, ActiveDay, Branch, Status_Active
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Cust_Param_Bonus'
Execute sp_executesql @SQL



Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Cust_option
INSERT Cust_option (Card_Type, Amount, Multiple, Expired_Name, Expired_Count, Active_Day, NewVal_Confirm, NewVal_Amount)
SELECT Card_Type, Amount, Multiple, Expired_Name, Expired_Count, Active_Day, NewVal_Confirm, NewVal_Amount
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Cust_option'
Execute sp_executesql @SQL

--truncate
Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE customer_master_member
INSERT customer_master_member(Cust_Nr, Cust_ID, Cust_Reg_Area, Cust_Name, Cust_Birth_Place, Cust_Birth_Date, Cust_Gender, Cust_Status, Cust_Address_Street1, 
                      Cust_Address_Street2, Cust_Address_City, Cust_Address_Zip, Cust_Phone_Area, Cust_Phone_Home, Cust_Phone_GSM, Cust_Phone_CDMA, 
                      Cust_Email, Cust_Registered, User_ID, Reg_Time, Data_Status, Agama, Pekerjaan, Ext1, Ext2, rowguid)
SELECT Cust_Nr, Cust_ID, Cust_Reg_Area, Cust_Name, Cust_Birth_Place, Cust_Birth_Date, Cust_Gender, Cust_Status, Cust_Address_Street1, 
                      Cust_Address_Street2, Cust_Address_City, Cust_Address_Zip, Cust_Phone_Area, Cust_Phone_Home, Cust_Phone_GSM, Cust_Phone_CDMA, 
                      Cust_Email, Cust_Registered, User_ID, Reg_Time, Data_Status, Agama, Pekerjaan, Ext1, Ext2, rowguid
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.customer_master_member'
Execute sp_executesql @SQL 

---spp_DownLoadCpoint '[192.168.1.23]',POS_SERVER_test

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Voucher_Master
INSERT Voucher_Master ( Type, Name, PLU, Price)
SELECT  Type, Name, PLU, Price
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Voucher_Master'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Card_Promotion
INSERT Card_Promotion ( Card_Nr, Card_Nr_Promo, Card_Promo_Id, Card_Expired_date, Card_Activate_Date, Card_Status, User_Id_Activate, rowguid)
SELECT  Card_Nr, Card_Nr_Promo, Card_Promo_Id, Card_Expired_date, Card_Activate_Date, Card_Status, User_Id_Activate, rowguid
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Card_Promotion'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Card_Promotion_Name
INSERT Card_Promotion_Name (Card_Promo_Id, Card_Promo_Name, Point_Bonus, Operator, Start_Promo_Date, End_Promo_Date, Data_Reg_Date, Data_Update_Date, 
                      Data_Status, User_Reg, User_Update, Card_Promo_Name_Long)
SELECT Card_Promo_Id, Card_Promo_Name, Point_Bonus, Operator, Start_Promo_Date, End_Promo_Date, Data_Reg_Date, Data_Update_Date, 
                      Data_Status, User_Reg, User_Update, Card_Promo_Name_Long
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Card_Promotion_Name'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Item_Promotion
INSERT  Item_Promotion(PLU, Item_Description, Class, Dp2, Disc, AppName)
SELECT PLU, Item_Description, Class, Dp2, Disc, AppName
FROM '+ rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Item_Promotion'
Execute sp_executesql @SQL

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Transaction_H_MemberCard'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Transaction_H'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Transaction_D_MemberCard'
Execute sp_executesql @SQL 

Set @SQL = '
Set Quoted_Identifier off
SET XACT_ABORT ON
TRUNCATE TABLE Customer_Transaction_D'
Execute sp_executesql @SQL

---spp_DownLoadCpoint '[192.168.1.23]',POS_SERVER_test




GO


