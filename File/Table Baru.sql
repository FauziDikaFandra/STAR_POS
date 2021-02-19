USE [POS_SERVER]
GO

/****** Object:  Table [dbo].[SPG]    Script Date: 10/04/2012 17:03:21 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[SPG](
	[spg_id] [varchar](6) NULL,
	[spg_name] [varchar](30) NULL,
	[spg_brand] [varchar](30) NULL,
	[spg_sbu] [varchar](20) NULL,
	[no_serial] [Int] NULL,
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


USE [POS_SERVER]
GO

/****** Object:  Table [dbo].[INFORMASI]    Script Date: 10/04/2012 17:03:28 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Informasi](
	[pesan1] [varchar](70) NULL,
	[pesan2] [varchar](70) NULL,
	[pesan3] [varchar](70) NULL,
	[pesan4] [varchar](70) NULL,
	[pesan5] [varchar](70) NULL,
	[pesan6] [varchar](70) NULL,
	[pesan7] [varchar](70) NULL,
	[pesan8] [varchar](70) NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

USE [POS_SERVER]
GO

/****** Object:  Table [dbo].[Cust_Param_Bonus]    Script Date: 10/10/2012 10:12:40 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Cust_Param_Bonus](
	[Jenis_Kartu] [varchar](2) NOT NULL,
	[Event_Name] [varchar](50) NOT NULL,
	[Point] [smallint] NOT NULL,
	[Start] [smalldatetime] NOT NULL,
	[Finish] [smalldatetime] NOT NULL,
	[ActiveDay] [varchar](7) NOT NULL,
	[Branch] [varchar](4) NOT NULL,
	[Status_Active] [varchar](1) NOT NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

USE [POS_SERVER]
GO

/****** Object:  Table [dbo].[Cust_Option]    Script Date: 10/10/2012 10:13:35 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Cust_Option](
	[Card_Type] [varchar](2) NOT NULL,
	[Amount] [money] NOT NULL,
	[Multiple] [varchar](1) NOT NULL,
	[Expired_Name] [varchar](2) NULL,
	[Expired_Count] [varchar](4) NULL,
	[Active_Day] [varchar](7) NULL,
	[NewVal_Confirm] [varchar](1) NULL,
	[NewVal_Amount] [money] NULL,
 CONSTRAINT [PK_Cust_Option] PRIMARY KEY CLUSTERED 
(
	[Card_Type] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

USE [POS_SERVER]
GO

/****** Object:  Table [dbo].[PROMO_HDR]    Script Date: 10/11/2012 16:53:20 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Promo_Hdr](
	[promo_id] [varchar](3) NOT NULL,
	[promo_name] [varchar](50) NULL,
	[start_date] [datetime] NULL,
	[end_date] [datetime] NULL,
	[min_purchase] [decimal](18, 0) NULL,
	[disc] [tinyint] NULL,
	[tipe] [tinyint] NULL,
 CONSTRAINT [PK_PROMO_HDR] PRIMARY KEY CLUSTERED 
(
	[promo_id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

USE [POS_SERVER]
GO

/****** Object:  Table [dbo].[PROMO_DTL]    Script Date: 10/16/2012 13:00:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Promo_Dtl](
	[promo_id] [varchar](3) NOT NULL,
	[PLU] [char](18) NOT NULL,
	[description] [char](20) NULL,
 CONSTRAINT [PK_PROMO_DTL] PRIMARY KEY CLUSTERED 
(
	[promo_id] ASC,
	[PLU] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO




