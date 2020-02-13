USE [LAB]
GO

/****** Object:  Table [dbo].[SAPMM-WM]    Script Date: 13/2/2563 13:43:01 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[SAPMM-WM](
	[Material Type] [nvarchar](10) NULL,
	[Material Type Description] [nvarchar](255) NULL,
	[Material Group] [nvarchar](10) NULL,
	[Matl Grp Desc#] [nvarchar](255) NULL,
	[Material] [nvarchar](12) NULL,
	[Description] [nvarchar](255) NULL,
	[Posting Date] [datetime] NULL,
	[ได้รับมาจาก] [nvarchar](255) NULL,
	[จ่ายไปให้] [nvarchar](255) NULL,
	[Movement type] [nvarchar](10) NULL,
	[Mvt Type Text] [nvarchar](255) NULL,
	[Batch] [nvarchar](100) NULL,
	[MFG Date] [datetime] NULL,
	[Manufacturer Batch] [nvarchar](100) NULL,
	[Manufacturer] [nvarchar](100) NULL,
	[Manufacturer Name] [nvarchar](255) NULL,
	[Vendor] [nvarchar](100) NULL,
	[Vendor Name] [nvarchar](255) NULL,
	[Sold-to] [nvarchar](20) NULL,
	[Sold-to Name] [nvarchar](255) NULL,
	[Sold-to Address] [nvarchar](255) NULL,
	[Sold-to Province] [nvarchar](100) NULL,
	[Ship-to] [nvarchar](20) NULL,
	[Ship-to Name] [nvarchar](255) NULL,
	[Ship-to Address] [nvarchar](255) NULL,
	[Ship-to Province] [nvarchar](100) NULL,
	[Customer Group 1] [nvarchar](5) NULL,
	[Customer Group 1 - Desc#] [nvarchar](100) NULL,
	[Customer Group 2] [nvarchar](5) NULL,
	[Customer Group 2 - Desc#] [nvarchar](100) NULL,
	[Customer Group 3] [nvarchar](5) NULL,
	[Customer Group 3 - Desc#] [nvarchar](100) NULL,
	[FG material] [nvarchar](20) NULL,
	[FG Material Description] [nvarchar](255) NULL,
	[FG Batch] [nvarchar](100) NULL,
	[Cost Center] [nvarchar](5) NULL,
	[Cost Center Description] [nvarchar](255) NULL,
	[Plant] [nvarchar](10) NULL,
	[PlantXXX] [nvarchar](10) NULL,
	[Storage Loc#] [nvarchar](10) NULL,
	[Dest# Plant] [nvarchar](10) NULL,
	[Dest# Sloc] [nvarchar](10) NULL,
	[ยอดยกมา] [float] NULL,
	[ปริมาณรับ] [float] NULL,
	[ปริมาณจ่าย] [float] NULL,
	[ปริมาณคงเหลือ] [float] NULL,
	[Unit] [nvarchar](20) NULL,
	[หมายเหตุ] [nvarchar](255) NULL,
	[Entered on] [datetime] NULL,
	[Entered at] [datetime] NULL,
	[Material Doc#] [nvarchar](20) NULL,
	[Mat# Doc# Year] [nvarchar](10) NULL,
	[Mat# Doc#Item] [nvarchar](5) NULL,
	[impdate] [datetime] NULL
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[SAPMM-WM] ADD  CONSTRAINT [DF_SAPMM-WM_impdate]  DEFAULT (getdate()) FOR [impdate]
GO

