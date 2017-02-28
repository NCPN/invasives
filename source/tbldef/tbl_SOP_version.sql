CREATE TABLE [tbl_SOP_version] (
  [version_key_number] LONG  CONSTRAINT [{43E11F6C-19F2-4D22-AE94-66B75DFFE700}] REFERENCES [tbl_master_version] ([version_key_number]),
  [SOP_number] LONG ,
  [SOP_version_number] DECIMAL (18, 2),
  [active_flag] BIT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([version_key_number], [SOP_number])
)
