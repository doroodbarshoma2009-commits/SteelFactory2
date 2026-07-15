/*
  SteelFactory2 — SQL Server Express Schema
  اجرا در SSMS روی سرور
*/
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = N'SteelFactory')
BEGIN
    CREATE DATABASE SteelFactory;
END
GO

USE SteelFactory;
GO

-- وضعیت کامل برنامه (معادل slab_db.json)
IF OBJECT_ID('dbo.app_state', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.app_state (
        id          INT NOT NULL PRIMARY KEY DEFAULT 1,
        data_json   NVARCHAR(MAX) NOT NULL,
        checksum    CHAR(64) NOT NULL,
        updated_at  DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        updated_by  NVARCHAR(100) NULL,
        client_id   NVARCHAR(100) NULL,
        CONSTRAINT CK_app_state_single CHECK (id = 1)
    );
END
GO

-- لاگ عملیات (برای مدیر و سرور)
IF OBJECT_ID('dbo.audit_log', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.audit_log (
        id          BIGINT IDENTITY(1,1) PRIMARY KEY,
        action      NVARCHAR(50) NOT NULL,
        user_name   NVARCHAR(100) NULL,
        client_id   NVARCHAR(100) NULL,
        details     NVARCHAR(MAX) NULL,
        created_at  DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME()
    );
    CREATE INDEX IX_audit_log_created ON dbo.audit_log(created_at DESC);
END
GO

-- بک‌آپ‌های خودکار
IF OBJECT_ID('dbo.backups', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.backups (
        id          BIGINT IDENTITY(1,1) PRIMARY KEY,
        data_json   NVARCHAR(MAX) NOT NULL,
        checksum    CHAR(64) NOT NULL,
        created_at  DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        created_by  NVARCHAR(100) NULL,
        client_id   NVARCHAR(100) NULL
    );
    CREATE INDEX IX_backups_created ON dbo.backups(created_at DESC);
END
GO

-- نشست‌های کلاینت
IF OBJECT_ID('dbo.client_sessions', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.client_sessions (
        id          BIGINT IDENTITY(1,1) PRIMARY KEY,
        client_id   NVARCHAR(100) NOT NULL,
        client_name NVARCHAR(200) NULL,
        ip_address  NVARCHAR(45) NULL,
        last_seen   DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        is_online   BIT NOT NULL DEFAULT 1
    );
    CREATE UNIQUE INDEX UX_client_sessions ON dbo.client_sessions(client_id);
END
GO

PRINT N'SteelFactory schema ready.';
