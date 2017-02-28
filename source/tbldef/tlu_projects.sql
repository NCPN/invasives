CREATE TABLE [tlu_projects] (
  [project_ID] LONG  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [project_name] VARCHAR (50),
  [project_linkfile] VARCHAR (255),
  [project_include_flag] BIT ,
  [project_comments] LONGTEXT 
)
