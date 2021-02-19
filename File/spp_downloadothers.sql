USE [POS_LOCAL]
GO

/****** Object:  StoredProcedure [dbo].[spp_DownloadOthers]    Script Date: 10/16/2012 13:54:36 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[spp_DownloadOthers]	
@ServerName char(20)
,@DBServer char(20)

AS
declare @SQL nvarchar(2000)

Set @SQL = '
Set Quoted_Identifier off 
SET XACT_ABORT ON
TRUNCATE TABLE SPG
insert into spg(spg_id, spg_name, spg_brand, spg_sbu, no_serial)
SELECT spg_id, spg_name, spg_brand, spg_sbu, no_serial
FROM '+rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.SPG'
Execute sp_executesql @SQL
print '1'

Set @SQL = '
Set Quoted_Identifier off 
SET XACT_ABORT ON
TRUNCATE TABLE INFORMASI
insert into INFORMASI(pesan1, pesan2, pesan3, pesan4, pesan5, pesan6, pesan7, pesan8)
SELECT   pesan1, pesan2, pesan3, pesan4, pesan5, pesan6, pesan7, pesan8
FROM '+rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.INFORMASI'
Execute sp_executesql @SQL
print '2'

Set @SQL = '
Set Quoted_Identifier off 
SET XACT_ABORT ON
TRUNCATE TABLE Key_Map
insert into Key_Map(form, menu, keyCode)
SELECT form, menu, keyCode
FROM '+rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.Key_Map'
Execute sp_executesql @SQL
print '3'

Set @SQL = '
Set Quoted_Identifier off 
SET XACT_ABORT ON
TRUNCATE TABLE PROMO_HDR
insert into Promo_Hdr(promo_id, promo_name, start_date, end_date, min_purchase, disc, tipe)
SELECT promo_id, promo_name, start_date, end_date, min_purchase, disc, tipe
FROM '+rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.PROMO_HDR'
Execute sp_executesql @SQL


Set @SQL = '
Set Quoted_Identifier off 
SET XACT_ABORT ON
TRUNCATE TABLE PROMO_DTL
Insert into Promo_Dtl(promo_id, PLU, description)
SELECT promo_id, PLU, description
FROM '+rtrim(@Servername)+'.'+rtrim(@DBServer)+'.dbo.PROMO_DTL'
Execute sp_executesql @SQL


GO


