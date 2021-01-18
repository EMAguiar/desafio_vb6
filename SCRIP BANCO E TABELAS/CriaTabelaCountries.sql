USE [DESAFIO]
GO

/****** Object:  Table [dbo].[FullCountryInfoAllCountries]    Script Date: 18/01/2021 05:50:54 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[FullCountryInfoAllCountries](
	[ISOCodeC] [char](2) NOT NULL,
	[Name] [nchar](100) NOT NULL,
	[CapitalCity] [nchar](100) NOT NULL,
	[PhoneCode] [nchar](3) NOT NULL,
	[ContinentCode] [nchar](2) NOT NULL,
	[CurrencyISOCode] [nchar](3) NOT NULL,
	[CountryFlag] [nchar](100) NOT NULL,
 CONSTRAINT [PK_FullCountryInfoAllCountries_1] PRIMARY KEY CLUSTERED 
(
	[ISOCodeC] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

