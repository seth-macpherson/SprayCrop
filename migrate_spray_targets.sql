USE agspray_dev
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE TABLE [dbo].[SprayRecordTargets](
	[SprayRecordID] [int] NOT NULL,
	[TargetID] [int] NOT NULL,
 CONSTRAINT [PK_SprayRecordTargets] PRIMARY KEY CLUSTERED 
(
	[SprayRecordID] ASC,
	[TargetID] ASC
)WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]


-- BEGIN TRAN

	INSERT INTO SprayRecordTargets (
		SprayRecordID,
		TargetID
	)
	SELECT SprayRecordID, TargetID
	FROM SprayRecord

SELECT *
FROM SprayRecordTargets
	
	
DROP TABLE SprayRecordTargets
	

-- ROLLBACK