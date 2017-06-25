CREATE DATABASE /*!32312 IF NOT EXISTS*/ `software_monitor` /*!40100 DEFAULT CHARACTER SET utf8mb4 */;
USE `software_monitor`;

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_AUTO_VALUE_ON_ZERO,NO_ENGINE_SUBSTITUTION,NO_ZERO_DATE,NO_ZERO_IN_DATE,STRICT_ALL_TABLES' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `device_details`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `device_details`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE IF NOT EXISTS `device_details` (
  `DeviceId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `DeviceParrentName` VARCHAR(60) DEFAULT NULL,
  `DeviceName` VARCHAR(60) NOT NULL,
  `DeviceOSdetails` JSON DEFAULT NULL,
  `DeviceHardwareDetails` JSON DEFAULT NULL,
  `MostRecentEvaluationId` MEDIUMINT(8) UNSIGNED DEFAULT NULL,
  `CountedEvaluations` MEDIUMINT(8) UNSIGNED DEFAULT NULL,
  `MostRecentSecurityEvaluationId` MEDIUMINT(8) UNSIGNED DEFAULT NULL,
  `CountedSecurityEvaluations` MEDIUMINT(8) UNSIGNED DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`DeviceId`),
  UNIQUE KEY `ndx_dd_DeviceName_UNIQUE` (`DeviceName` ASC),
  KEY `ndx_dd_DeviceParrentName` (`DeviceParrentName`),
  KEY `FK_dd_MostRecentEvaluationId` (`MostRecentEvaluationId`),
  KEY `FK_dd_MostRecentSecurityEvaluationId` (`MostRecentSecurityEvaluationId`),
  CONSTRAINT `FK_dd_MostRecentEvaluationId` 
    FOREIGN KEY (`MostRecentEvaluationId`) 
    REFERENCES `evaluation_headers` (`EvaluationId`) 
    ON DELETE RESTRICT
    ON UPDATE CASCADE,
  CONSTRAINT `FK_dd_MostRecentSecurityEvaluationId` 
    FOREIGN KEY (`MostRecentSecurityEvaluationId`) 
    REFERENCES `security_evaluation_headers` (`SecurityEvaluationId`) 
    ON DELETE RESTRICT
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT 'List of devices in scope';
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `device_volumes`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `device_volumes`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `device_volumes` (
  `DeviceVolumeId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `VolumeSerialNumber` VARCHAR(60) NOT NULL,
  `DetailedInformation` JSON NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`DeviceVolumeId`),
  UNIQUE KEY `ndx_dd_VolumeSerialNumber_UNIQUE` (`VolumeSerialNumber` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `in_windows_software_installed`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `in_windows_software_installed`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `in_windows_software_installed` (
  `EvaluationTimestamp` DATETIME NOT NULL,
  `HostName` VARCHAR(64) NOT NULL,
  `PublisherName` VARCHAR(80) DEFAULT NULL,
  `SoftwareName` VARCHAR(254) NOT NULL,
  `FullVersion` VARCHAR(30) DEFAULT NULL,
  `InstallationDate` DATE DEFAULT NULL,
  `OtherInfo` JSON DEFAULT NULL,
  `RegistryHive` ENUM('HKEY_LOCAL_MACHINE', 'HKEY_CURRENT_USER', 'Unknown') NOT NULL DEFAULT 'Unknown',
  `RegistrySubKey` VARCHAR(200) NOT NULL,
  `Bits32Or64` ENUM('0', '32', '64') NOT NULL DEFAULT '0',
  PRIMARY KEY(`HostName`, `RegistryHive`, `RegistrySubKey`, `Bits32Or64`),
  KEY `ndx_iwsi_HostName` (`HostName`),
  KEY `ndx_iwsi_PublisherName` (`PublisherName`),
  KEY `ndx_iwsi_SoftwareName` (`SoftwareName`),
  KEY `ndx_iwsi_FullVersion` (`FullVersion`),
  KEY `ndx_iwsi_InstallationDate` (`InstallationDate`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `in_windows_software_portable`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `in_windows_software_portable`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `in_windows_software_portable` (
  `EvaluationTimestamp` DATETIME NOT NULL,
  `VolumeSerialNumber` VARCHAR(50) NOT NULL,
  `FileNameSearched` VARCHAR(100) NOT NULL,
  `MethodToFind` ENUM('Aproximate', 'Exact') NOT NULL,
  `FilePathFound` VARCHAR(255) NOT NULL,
  `FileNameFound` VARCHAR(100) NOT NULL,
  `FileDateCreated` DATETIME NOT NULL,
  `FileDateLastModified` DATETIME NOT NULL,
  `FileVersionFound` VARCHAR(30) DEFAULT NULL,
  `FileSizeFound` MEDIUMINT(8) UNSIGNED DEFAULT NULL,
  `FilesCheckedForMatchUntilFound` MEDIUMINT(8) UNSIGNED NOT NULL,
  PRIMARY KEY(`VolumeSerialNumber`, `FileNameSearched`, `FilePathFound`, `FileNameFound`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `in_windows_security_risk_components`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `in_windows_security_risk_components`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `in_windows_security_risk_components` (
  `EvaluationTimestamp` DATETIME NOT NULL,
  `VolumeSerialNumber` VARCHAR(50) NOT NULL,
  `FileNameSearched` VARCHAR(100) NOT NULL,
  `MethodToFind` ENUM('Aproximate', 'Exact') NOT NULL,
  `FilePathFound` VARCHAR(255) NOT NULL,
  `FileNameFound` VARCHAR(100) NOT NULL,
  `FileDateCreated` DATETIME NOT NULL,
  `FileDateLastModified` DATETIME NOT NULL,
  `FileVersionFound` VARCHAR(30) DEFAULT NULL,
  `FileSizeFound` mediumint(8) unsigned DEFAULT NULL,
  `FilesCheckedForMatchUntilFound` mediumint(8) unsigned NOT NULL,
  PRIMARY KEY(`VolumeSerialNumber`, `FileNameSearched`, `FilePathFound`, `FileNameFound`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `publisher_details`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `publisher_details`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `publisher_details` (
  `PublisherId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `PublisherName` VARCHAR(80) NOT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`PublisherId`),
  UNIQUE KEY `ndx_pd_PublisherName_UNIQUE` (`PublisherName` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `publisher_known`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `publisher_known`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `publisher_known` (
  `PublisherName` VARCHAR(80) NOT NULL,
  `PublisherMainWebsite` VARCHAR(250) NULL DEFAULT NULL,
  `PublisherExtendedInformation` JSON NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`PublisherName`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `software_details`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `software_details`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `software_details` (
  `SoftwareId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `SoftwareName` VARCHAR(254) NOT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`SoftwareId`),
  UNIQUE KEY `ndx_sd_SoftwareName_UNIQUE` (`SoftwareName` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `software_known`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `software_known`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `software_known` (
  `SoftwareName` VARCHAR(80) NOT NULL,
  `SoftwareDescription` TEXT NULL DEFAULT NULL,
  `SoftwareMainWebsite` VARCHAR(250) NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`SoftwareName`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `software_files`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `software_files`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE IF NOT EXISTS `software_files` (
  `SoftwareFileName` VARCHAR(100) NOT NULL,
  `SoftwareFileVersionFirst` VARCHAR(30) NOT NULL,
  `SoftwareFileVersionPiecesFirst` JSON GENERATED ALWAYS AS (CONCAT('{ "Major": ', (CASE WHEN (`SoftwareFileVersionFirst` IS NULL) THEN 0 ELSE CAST(REPLACE((CASE WHEN (LOCATE(".", `SoftwareFileVersionFirst`) = 0) THEN `SoftwareFileVersionFirst` ELSE SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 1) END), "v", "") AS UNSIGNED) END), ', "Minor": ', (CASE WHEN (`SoftwareFileVersionFirst` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionFirst`) = 0) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 2), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 1), '.'), '') END) AS UNSIGNED) END), ', "Build": ',(CASE WHEN (`SoftwareFileVersionFirst` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionFirst`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`SoftwareFileVersionFirst`) - CHAR_LENGTH(REPLACE(`SoftwareFileVersionFirst`, '.', ''))) < 2) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 3), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 2), '.'), '') END) AS UNSIGNED) END), ', "Revision": ', (CASE WHEN (`SoftwareFileVersionFirst` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionFirst`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`SoftwareFileVersionFirst`) - CHAR_LENGTH(REPLACE(`SoftwareFileVersionFirst`, '.', ''))) < 3) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 4), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 3), '.'), '') END) AS UNSIGNED) END), ' }')) STORED,
  `SoftwareFileVersionNumericFirst` DECIMAL(25,7) ZEROFILL GENERATED ALWAYS AS (CASE WHEN (`SoftwareFileVersionFirst` IS NULL) THEN NULL ELSE CAST( ( CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesFirst`, '$.Major') * POW(10, 14)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesFirst`, '$.Minor') * POW(10, 7)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesFirst`, '$.Build') * POW(10, 0)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesFirst`, '$.Revision') / POW(10, 7)) AS DECIMAL(8, 7)) ) AS DECIMAL(30, 7)) END) STORED,
  `SoftwareFileVersionLast` VARCHAR(30) NOT NULL,
  `SoftwareFileVersionPiecesLast` JSON GENERATED ALWAYS AS (CONCAT('{ "Major": ', (CASE WHEN (`SoftwareFileVersionLast` IS NULL) THEN NULL ELSE CAST(REPLACE((CASE WHEN (LOCATE(".", `SoftwareFileVersionLast`) = 0) THEN `SoftwareFileVersionLast` ELSE SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 1) END), "v", "") AS UNSIGNED) END), ', "Minor": ', (CASE WHEN (`SoftwareFileVersionLast` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionLast`) = 0) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 2), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 1), '.'), '') END) AS UNSIGNED) END), ', "Build": ',(CASE WHEN (`SoftwareFileVersionLast` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionLast`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`SoftwareFileVersionLast`) - CHAR_LENGTH(REPLACE(`SoftwareFileVersionLast`, '.', ''))) < 2) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 3), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 2), '.'), '') END) AS UNSIGNED) END), ', "Revision": ', (CASE WHEN (`SoftwareFileVersionLast` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionLast`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`SoftwareFileVersionLast`) - CHAR_LENGTH(REPLACE(`SoftwareFileVersionLast`, '.', ''))) < 3) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 4), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 3), '.'), '') END) AS UNSIGNED) END), ' }')) STORED,
  `SoftwareFileVersionNumericLast` DECIMAL(25,7) ZEROFILL GENERATED ALWAYS AS (CASE WHEN (`SoftwareFileVersionLast` IS NULL) THEN NULL ELSE CAST( ( CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesLast`, '$.Major') * POW(10, 14)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesLast`, '$.Minor') * POW(10, 7)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesLast`, '$.Build') * POW(10, 0)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesLast`, '$.Revision') / POW(10, 7)) AS DECIMAL(8, 7)) ) AS DECIMAL(30, 7)) END) STORED,
  `SoftwareName` VARCHAR(80) NOT NULL,
  `PublisherName` VARCHAR(80) NOT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`SoftwareFileName`, `SoftwareFileVersionFirst`, `SoftwareFileVersionLast`),
  KEY `ndx_sf_SoftwareName` (`SoftwareName`),
  KEY `ndx_sf_PublisherName` (`PublisherName`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT 'Known files w/ their standardized Software & relevant Publisher';
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `version_details`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `version_details`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `version_details` (
    `VersionId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
    `FullVersion` VARCHAR(30) NOT NULL,
    `FullVersionParts` JSON GENERATED ALWAYS AS (CONCAT('{ "Major": ', (CASE WHEN (`FullVersion` IS NULL) THEN NULL ELSE CAST(REPLACE((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN `FullVersion` ELSE SUBSTRING_INDEX(`FullVersion`, '.', 1) END), "v", "") AS UNSIGNED) END), ', "Minor": ', (CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 2), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 1), '.'), '') END) AS UNSIGNED) END), ', "Build": ',(CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`FullVersion`) - CHAR_LENGTH(REPLACE(`FullVersion`, '.', ''))) < 2) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 3), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 2), '.'), '') END) AS UNSIGNED) END), ', "Revision": ', (CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`FullVersion`) - CHAR_LENGTH(REPLACE(`FullVersion`, '.', ''))) < 3) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 4), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 3), '.'), '') END) AS UNSIGNED) END), ' }')) STORED,
    `FullVersionNumeric` DECIMAL(30,7) ZEROFILL GENERATED ALWAYS AS (CASE WHEN (`FullVersion` IS NULL) THEN NULL ELSE CAST( ( CAST((JSON_EXTRACT(`FullVersionParts`, '$.Major') * POW(10, 14)) AS UNSIGNED) + CAST((JSON_EXTRACT(`FullVersionParts`, '$.Minor') * POW(10, 7)) AS UNSIGNED) + CAST((JSON_EXTRACT(`FullVersionParts`, '$.Build') * POW(10, 0)) AS UNSIGNED) + CAST((JSON_EXTRACT(`FullVersionParts`, '$.Revision') / POW(10, 7)) AS DECIMAL(8, 7)) ) AS DECIMAL(30, 7)) END) STORED,
    PRIMARY KEY(`VersionId`),
    UNIQUE KEY `ndx_vd_FullVersion_UNIQUE` (`FullVersion` ASC),
    KEY `ndx_vd_FullVersionNumeric` (`FullVersionNumeric`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/* Evaluation structures to support traceability for Software */

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `evaluation_headers`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `evaluation_headers`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `evaluation_headers` (
  `EvaluationId` MEDIUMINT(8) UNSIGNED NOT NULL AUTO_INCREMENT,
  `DeviceId` SMALLINT(5) UNSIGNED NOT NULL,
  `EvaluationTimestamp` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `DateOfGatheringTimestampFirst` DATETIME DEFAULT NULL,
  `DateOfGatheringTimestampLast` DATETIME DEFAULT NULL,
  `GatheringDuration` VARCHAR(50) GENERATED ALWAYS AS (TRIM(CAST((CASE WHEN ((`DateOfGatheringTimestampFirst` IS NOT NULL) AND (`DateOfGatheringTimestampLast` IS NOT NULL)) THEN TIMEDIFF(`DateOfGatheringTimestampLast`, `DateOfGatheringTimestampFirst`) ELSE NULL END) AS CHAR(50)))) STORED,
  PRIMARY KEY (`EvaluationId`),
  CONSTRAINT `FK_eh_DeviceId`
    FOREIGN KEY (`DeviceId`)
    REFERENCES `device_details` (`DeviceId`)
    ON DELETE RESTRICT
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `evaluation_lines`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `evaluation_lines`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `evaluation_lines` (
  `EvaluationId` MEDIUMINT(8) UNSIGNED NOT NULL,
  `PublisherId` SMALLINT(5) UNSIGNED NOT NULL,
  `SoftwareId` SMALLINT(5) UNSIGNED NOT NULL,
  `VersionId` SMALLINT(5) UNSIGNED NOT NULL,
  `InstallationDate` DATE DEFAULT NULL,
  `Folders` TEXT DEFAULT NULL,
  CONSTRAINT `FK_el_EvaluationId`
    FOREIGN KEY (`EvaluationId`)
    REFERENCES `evaluation_headers` (`EvaluationId`)
    ON DELETE RESTRICT
    ON UPDATE CASCADE,
  CONSTRAINT `FK_el_PublisherId`
    FOREIGN KEY (`PublisherId`)
    REFERENCES `publisher_details` (`PublisherId`)
    ON DELETE RESTRICT
    ON UPDATE CASCADE,
  CONSTRAINT `FK_el_SoftwareId`
    FOREIGN KEY (`SoftwareId`)
    REFERENCES `software_details` (`SoftwareId`)
    ON DELETE RESTRICT
    ON UPDATE CASCADE,
  CONSTRAINT `FK_el_VersionId`
    FOREIGN KEY (`VersionId`)
    REFERENCES `version_details` (`VersionId`)
    ON DELETE RESTRICT
    ON UPDATE CASCADE,
  UNIQUE KEY `ndx_el_MultipleFields_UNIQUE` (`EvaluationId` ASC, `PublisherId` ASC, `SoftwareId` ASC, `VersionId` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `security_evaluation_headers`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `security_evaluation_headers`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `security_evaluation_headers` (
  `SecurityEvaluationId` MEDIUMINT(8) UNSIGNED NOT NULL AUTO_INCREMENT,
  `DeviceId` SMALLINT(5) UNSIGNED NOT NULL,
  `EvaluationTimestamp` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `DateOfGatheringTimestampFirst` DATETIME DEFAULT NULL,
  `DateOfGatheringTimestampLast` DATETIME DEFAULT NULL,
  `GatheringDuration` VARCHAR(50) GENERATED ALWAYS AS (TRIM(CAST((CASE WHEN ((`DateOfGatheringTimestampFirst` IS NOT NULL) AND (`DateOfGatheringTimestampLast` IS NOT NULL)) THEN TIMEDIFF(`DateOfGatheringTimestampLast`, `DateOfGatheringTimestampFirst`) ELSE NULL END) AS CHAR(50)))) STORED,
  PRIMARY KEY (`SecurityEvaluationId`),
  CONSTRAINT `FK_seh_DeviceId`
    FOREIGN KEY (`DeviceId`)
    REFERENCES `device_details` (`DeviceId`)
    ON DELETE RESTRICT
    ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Table structure for table `security_evaluation_lines`
/*--------------------------------------------------------------------------------------------------------------------*/
DROP TABLE IF EXISTS `security_evaluation_lines`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE IF NOT EXISTS `security_evaluation_lines` (
  `SecurityEvaluationId` MEDIUMINT(8) UNSIGNED NOT NULL,
  `PublisherId` SMALLINT(5) UNSIGNED NOT NULL,
  `SoftwareId` SMALLINT(5) UNSIGNED NOT NULL,
  `VersionId` SMALLINT(5) UNSIGNED NOT NULL,
  `InstallationDate` DATE DEFAULT NULL,
  `Folders` TEXT DEFAULT NULL,
  CONSTRAINT `FK_sel_SecurityEvaluationId`
    FOREIGN KEY (`SecurityEvaluationId`)
    REFERENCES `security_evaluation_headers` (`SecurityEvaluationId`)
    ON DELETE RESTRICT
    ON UPDATE CASCADE,
  CONSTRAINT `FK_sel_PublisherId`
    FOREIGN KEY (`PublisherId`)
    REFERENCES `publisher_details` (`PublisherId`)
    ON DELETE RESTRICT
    ON UPDATE CASCADE,
  CONSTRAINT `FK_sel_SoftwareId`
    FOREIGN KEY (`SoftwareId`)
    REFERENCES `software_details` (`SoftwareId`)
    ON DELETE RESTRICT
    ON UPDATE CASCADE,
  CONSTRAINT `FK_sel_VersionId`
    FOREIGN KEY (`VersionId`)
    REFERENCES `version_details` (`VersionId`)
    ON DELETE RESTRICT
    ON UPDATE CASCADE,
  UNIQUE KEY `ndx_sel_MultipleFields_UNIQUE` (`SecurityEvaluationId` ASC, `PublisherId` ASC, `SoftwareId` ASC, `VersionId` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
/*!40101 SET character_set_client = @saved_cs_client */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Series of Views to provide a quick summary on various things
/*--------------------------------------------------------------------------------------------------------------------*/

/*--------------------------------------------------------------------------------------------------------------------*/
-- View structure for view `view__devices`
/*--------------------------------------------------------------------------------------------------------------------*/
/*!50001 DROP VIEW IF EXISTS `view__devices`*/;
/*!50001 SET @saved_cs_client          = @@character_set_client */;
/*!50001 SET @saved_cs_results         = @@character_set_results */;
/*!50001 SET @saved_col_connection     = @@collation_connection */;
/*!50001 SET character_set_client      = utf8 */;
/*!50001 SET character_set_results     = utf8 */;
/*!50001 SET collation_connection      = utf8_general_ci */;
/*!50001 CREATE ALGORITHM=UNDEFINED */
/*!50001 VIEW `view__devices` AS
SELECT
    `dd`.`DeviceId` AS `DeviceId`,
    `dd`.`DeviceParrentName` AS `DeviceParrentName`,
    `dd`.`DeviceName` AS `DeviceName`,
    (CASE WHEN (`dd`.`DeviceOSdetails` IS NOT NULL) THEN REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."Caption"'),'"','') ELSE REPLACE(JSON_EXTRACT(`dv`.`DetailedInformation`,'$."Description"'),'"','') END) AS `Caption`,
     (CASE WHEN (`dd`.`DeviceOSdetails` IS NOT NULL) THEN REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."OS Architecture"'),'"','') ELSE REPLACE(JSON_EXTRACT(`dv`.`DetailedInformation`,'$."File System"'),'"','') END) AS `Architecture`,
    REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."Version"'),'"','') AS `Version`,
    ROUND(CAST(REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."Total Visible Memory [MB]"'),'"','') AS UNSIGNED) / 1024, 0) AS `RAM [GB]`,
    REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."Current Time Zone Description"'),'"','') AS `Current Time Zone`,
    REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."OS Language Description"'),'"','') AS `Language`,
    REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."Locale Description"'),'"','') AS `Locale`,
    `dd`.`CountedEvaluations` AS `Number of Evaluations`,
    MAX(`eh`.`DateOfGatheringTimestampLast`) AS `Most Recent Evaluation Timestamp`,
    DATEDIFF(NOW(), MAX(`eh`.`DateOfGatheringTimestampLast`)) AS `Most Recent Evaluation Aging`,
    `dd`.`CountedSecurityEvaluations` AS `Number of Security Evaluations`,
    MAX(`seh`.`DateOfGatheringTimestampLast`) AS `Most Recent Security Evaluation Timestamp`,
    DATEDIFF(NOW(), MAX(`seh`.`DateOfGatheringTimestampLast`)) AS `Most Recent Security Evaluation Aging` 
FROM `device_details` `dd`
    LEFT JOIN `evaluation_headers` `eh` ON `dd`.`DeviceId` = `eh`.`DeviceId`
    LEFT JOIN `device_volumes` `dv` ON `dd`.`DeviceName` = `dv`.`VolumeSerialNumber`
    LEFT JOIN `security_evaluation_headers` `seh` ON `dd`.`DeviceId` = `seh`.`DeviceId`
GROUP BY `dd`.`DeviceId`, `dd`.`DeviceName`
ORDER BY `dd`.`DeviceParrentName`, `dd`.`DeviceName` */;
/*!50001 SET character_set_client      = @saved_cs_client */;
/*!50001 SET character_set_results     = @saved_cs_results */;
/*!50001 SET collation_connection      = @saved_col_connection */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- View structure for view `view__evaluations`
/*--------------------------------------------------------------------------------------------------------------------*/
/*!50001 DROP VIEW IF EXISTS `view__evaluations`*/;
/*!50001 SET @saved_cs_client          = @@character_set_client */;
/*!50001 SET @saved_cs_results         = @@character_set_results */;
/*!50001 SET @saved_col_connection     = @@collation_connection */;
/*!50001 SET character_set_client      = utf8 */;
/*!50001 SET character_set_results     = utf8 */;
/*!50001 SET collation_connection      = utf8_general_ci */;
/*!50001 CREATE ALGORITHM=UNDEFINED */
/*!50001 VIEW `view__evaluations` AS
    SELECT 
        `el`.`EvaluationId` AS `EvaluationId`,
        `dd`.`DeviceId` AS `DeviceId`,
        `dd`.`DeviceName` AS `DeviceName`,
        `pd`.`PublisherId` AS `PublisherId`,
        `pd`.`PublisherName` AS `PublisherName`,
        `sd`.`SoftwareId` AS `SoftwareId`,
        `sd`.`SoftwareName` AS `SoftwareName`,
        (CASE
            WHEN
                (`sd`.`SoftwareName` IN (
                    'Intel® Processor Graphics',
                    'Maxx Audio Installer',
                    'Microsoft redistributable runtime DLLs',
                    'Microsoft Visio Professional',
                    'Microsoft Visio Standard',
                    'Microsoft Visual C++ Additional Runtime',
                    'Microsoft Visual C++ Debug Runtime',
                    'Microsoft Visual C++ Minimum Runtime',
                    'Microsoft Visual C++ Redistributable',
                    'Microsoft Visual Studio Tools for Office Runtime',
                    'Mozilla Firefox',
                    'Office Click-to-Run Extensibility Component',
                    'Office Click-to-Run Licensing Component',
                    'Office Click-to-Run Localization Component',
                    'PHP',
                    'Security Update for Microsoft .NET Framework'
                    )
                )
                THEN
                    JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major')
            WHEN
                (`sd`.`SoftwareName` IN(
                    'Microsoft .NET Framework Multi-Targeting Pack'
                    )
                )
                THEN
                    CONCAT(JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major'), '.', JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Minor'))
            WHEN
                (`sd`.`SoftwareName` LIKE 'Intel® Management Engine Components')
                THEN
                    CONVERT( (CASE
                        WHEN (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 1) THEN 1
                        ELSE 'non 1'
                    END) USING UTF8MB4)
            WHEN
                (`sd`.`SoftwareName` IN (
                    'MySQL Documents',
                    'MySQL Examples and Samples',
                    'MySQL Server'
                    )
                )
                THEN
                    CONVERT( (CASE
                        WHEN (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 8) THEN 'dmr'
                        ELSE NULL
                    END) USING UTF8MB4)
            WHEN
                (`sd`.`SoftwareName` LIKE 'OpenSSL')
                THEN
                    CONVERT( (CASE
                        WHEN
                            ((JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 0)
                                AND (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Minor') = 9))
                        THEN
                            '0.9.x'
                        WHEN
                            ((JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 1)
                                AND (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Minor') = 0)
                                AND (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Build') = 0))
                        THEN
                            '1.0.0.x'
                        WHEN
                            ((JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 1)
                                AND (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Minor') = 0)
                                AND (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Build') = 1))
                        THEN
                            '1.0.1.x'
                        WHEN
                            ((JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 1)
                                AND (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Minor') = 0)
                                AND (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Build') = 2))
                        THEN
                            '1.0.2.x'
                        WHEN
                            ((JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 1)
                                AND (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Minor') = 1))
                        THEN
                            '1.1.x'
                        ELSE NULL
                    END) USING UTF8MB4)
            ELSE NULL
        END) AS `RelevantMajorVersion`,
        `vd`.`VersionId` AS `VersionId`,
        `vd`.`FullVersion` AS `FullVersion`,
        `vd`.`FullVersionNumeric` AS `FullVersionNumeric`
    FROM
        (((((`evaluation_lines` `el`
        JOIN `evaluation_headers` `eh` ON ((`el`.`EvaluationId` = `eh`.`EvaluationId`)))
        JOIN `device_details` `dd` ON (((`eh`.`EvaluationId` = `dd`.`MostRecentEvaluationId`)
            AND (`eh`.`DeviceId` = `dd`.`DeviceId`))))
        JOIN `publisher_details` `pd` ON ((`el`.`PublisherId` = `pd`.`PublisherId`)))
        JOIN `software_details` `sd` ON ((`el`.`SoftwareId` = `sd`.`SoftwareId`)))
        JOIN `version_details` `vd` ON ((`el`.`VersionId` = `vd`.`VersionId`))) */;
/*!50001 SET character_set_client      = @saved_cs_client */;
/*!50001 SET character_set_results     = @saved_cs_results */;
/*!50001 SET collation_connection      = @saved_col_connection */;

/*--------------------------------------------------------------------------------------------------------------------*/
-- View structure for view `view__version_assesment`
/*--------------------------------------------------------------------------------------------------------------------*/
/*!50001 DROP VIEW IF EXISTS `view__version_assesment`*/;
/*!50001 SET @saved_cs_client          = @@character_set_client */;
/*!50001 SET @saved_cs_results         = @@character_set_results */;
/*!50001 SET @saved_col_connection     = @@collation_connection */;
/*!50001 SET character_set_client      = utf8 */;
/*!50001 SET character_set_results     = utf8 */;
/*!50001 SET collation_connection      = utf8_general_ci */;
/*!50001 CREATE ALGORITHM=UNDEFINED */
/*!50001 VIEW `view__version_assesment` AS
SELECT
    `ve`.`PublisherName`,
    `ve`.`SoftwareName`,
    `ve`.`RelevantMajorVersion` AS `RMV`,
    GROUP_CONCAT(DISTINCT CONCAT(`ve`.`DeviceName`, '_', `ve`.`FullVersion`) ORDER BY `ve`.`DeviceName` SEPARATOR "
") AS `Version Details`,
    (CASE WHEN (SUM(CASE WHEN (`ve`.`DeviceId` = 3) THEN `ve`.`FullVersionNumeric` ELSE NULL END) IS NULL) THEN '---' WHEN SUM(CASE WHEN (`ve`.`DeviceId` = 3) THEN `ve`.`FullVersionNumeric` ELSE NULL END) = MAX(`ve`.`FullVersionNumeric`) THEN 'Ok' ELSE 'OLD' END) AS `PC_2016`,
    (CASE WHEN (SUM(CASE WHEN (`ve`.`DeviceId` = 2) THEN `ve`.`FullVersionNumeric` ELSE NULL END) IS NULL) THEN '---' WHEN SUM(CASE WHEN (`ve`.`DeviceId` = 2) THEN `ve`.`FullVersionNumeric` ELSE NULL END) = MAX(`ve`.`FullVersionNumeric`) THEN 'Ok' ELSE 'OLD' END) AS `PC_2014`,
    (CASE WHEN (SUM(CASE WHEN (`ve`.`DeviceId` = 1) THEN `ve`.`FullVersionNumeric` ELSE NULL END) IS NULL) THEN '---' WHEN SUM(CASE WHEN (`ve`.`DeviceId` = 1) THEN `ve`.`FullVersionNumeric` ELSE NULL END) = MAX(`ve`.`FullVersionNumeric`) THEN 'Ok' ELSE 'OLD' END) AS `PC_2013`,
    (CASE WHEN (SUM(CASE WHEN (`ve`.`DeviceId` = 5) THEN `ve`.`FullVersionNumeric` ELSE NULL END) IS NULL) THEN '---' WHEN SUM(CASE WHEN (`ve`.`DeviceId` = 5) THEN `ve`.`FullVersionNumeric` ELSE NULL END) = MAX(`ve`.`FullVersionNumeric`) THEN 'Ok' ELSE 'OLD' END) AS `Web_Server_H`,
    (CASE WHEN (SUM(CASE WHEN (`ve`.`DeviceId` = 4) THEN `ve`.`FullVersionNumeric` ELSE NULL END) IS NULL) THEN '---' WHEN SUM(CASE WHEN (`ve`.`DeviceId` = 4) THEN `ve`.`FullVersionNumeric` ELSE NULL END) = MAX(`ve`.`FullVersionNumeric`) THEN 'Ok' ELSE 'OLD' END) AS `Web_Server_M`,
    MAX(`ve`.`FullVersionNumeric`) AS `Newest`,
    MIN(`ve`.`FullVersionNumeric`) AS `Oldest`,
    GROUP_CONCAT(DISTINCT `ve`.`EvaluationId` SEPARATOR "; ") AS `Evaluations`,
    `ve`.`SoftwareId`,
    (CASE WHEN MIN(`ve`.`FullVersionNumeric`) = MAX(`ve`.`FullVersionNumeric`) THEN 'Everything up-to-date' ELSE 'Differences...' END) AS `Assesment`
FROM `view__evaluations` `ve`
WHERE (`ve`.`DeviceId` IN (1, 2, 3, 4, 5))
GROUP BY `ve`.`SoftwareName`, `ve`.`RelevantMajorVersion`
HAVING (`Assesment` = 'Differences...') */;
/*!50001 SET character_set_client      = @saved_cs_client */;
/*!50001 SET character_set_results     = @saved_cs_results */;
/*!50001 SET collation_connection      = @saved_col_connection */;

/* ------------------------------------------------------------------------------------------------------------------ */

/*--------------------------------------------------------------------------------------------------------------------*/
-- Strcture for Stored Procedure `pr_MatchLatestEvaluationForSoftwarePortrable`
/*--------------------------------------------------------------------------------------------------------------------*/
/*!50003 DROP PROCEDURE IF EXISTS `pr_MatchLatestEvaluationForSoftwarePortrable` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_AUTO_VALUE_ON_ZERO,STRICT_ALL_TABLES,NO_ZERO_IN_DATE,NO_ZERO_DATE,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER //
CREATE PROCEDURE `pr_MatchLatestEvaluationForSoftwarePortable`()
    NOT DETERMINISTIC
    READS SQL DATA
    SQL SECURITY DEFINER
    COMMENT 'Sets most recent evaluation to devices based on volumes'
BEGIN
    DECLARE v_DeviceId SMALLINT(5) UNSIGNED;
    DECLARE v_EvaluationId MEDIUMINT(8) UNSIGNED;
    DECLARE v_CountedEvaluations MEDIUMINT(8) UNSIGNED;
    DECLARE v_done INT DEFAULT 0;
    /* Reads existing AI columns to later evaluate 1 by 1 */
    DECLARE info_cursor CURSOR FOR SELECT `dd`.`DeviceId`, MAX(`eh`.`EvaluationId`), COUNT(DISTINCT `eh`.`EvaluationId`) FROM `in_windows_software_portable` `iwsp` INNER JOIN `device_details` `dd` ON `iwsp`.`VolumeSerialNumber` = `dd`.`DeviceName` INNER JOIN `evaluation_headers` `eh` ON `dd`.`DeviceId` = `eh`.`DeviceId` GROUP BY `dd`.`DeviceId`;
    DECLARE CONTINUE HANDLER FOR NOT FOUND SET v_done = 1;
    /* Evaluate current situation for every single relevant column and table */
    SET @dynamic_sql = "UPDATE `device_details` SET `MostRecentEvaluationId` = ?, `CountedEvaluations` = ?, `LastSeen` = `LastSeen` WHERE (`DeviceId` = ?);";
    PREPARE complete_sql FROM @dynamic_sql;
    OPEN info_cursor;
    REPEAT
        FETCH info_cursor INTO v_DeviceId, v_EvaluationId, v_CountedEvaluations;
        IF NOT v_done THEN
            SET @DeviceId = v_DeviceId;
            SET @EvaluationId = v_EvaluationId;
            SET @CountedEvaluations = v_CountedEvaluations;
            EXECUTE complete_sql USING @EvaluationId, @CountedEvaluations, @DeviceId;
        END IF;
    UNTIL v_done END REPEAT;
    CLOSE info_cursor;
    DEALLOCATE PREPARE complete_sql;
END//
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

/*--------------------------------------------------------------------------------------------------------------------*/
-- Strcture for Stored Procedure `pr_MatchLatestEvaluationForSecurityRiskComponents`
/*--------------------------------------------------------------------------------------------------------------------*/
/*!50003 DROP PROCEDURE IF EXISTS `pr_MatchLatestEvaluationForSecurityRiskComponents` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8 */ ;
/*!50003 SET character_set_results = utf8 */ ;
/*!50003 SET collation_connection  = utf8_general_ci */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'NO_AUTO_VALUE_ON_ZERO,STRICT_ALL_TABLES,NO_ZERO_IN_DATE,NO_ZERO_DATE,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
DELIMITER //
CREATE PROCEDURE `pr_MatchLatestEvaluationForSecurityRiskComponents`()
    NOT DETERMINISTIC
    READS SQL DATA
    SQL SECURITY DEFINER
    COMMENT 'Sets latest security risk assesment to devices based on volumes'
BEGIN
    DECLARE v_DeviceId SMALLINT(5) UNSIGNED;
    DECLARE v_SecurityEvaluationId MEDIUMINT(8) UNSIGNED;
    DECLARE v_CountedSecurityEvaluations MEDIUMINT(8) UNSIGNED;
    DECLARE v_done INT DEFAULT 0;
    /* Reads existing AI columns to later evaluate 1 by 1 */
    DECLARE info_cursor CURSOR FOR SELECT `dd`.`DeviceId`, MAX(`seh`.`SecurityEvaluationId`), COUNT(DISTINCT `seh`.`SecurityEvaluationId`) FROM `in_windows_security_risk_components` `iwsrc` INNER JOIN `device_details` `dd` ON `iwsrc`.`VolumeSerialNumber` = `dd`.`DeviceName` INNER JOIN `security_evaluation_headers` `seh` ON `dd`.`DeviceId` = `seh`.`DeviceId` GROUP BY `dd`.`DeviceId`;
    DECLARE CONTINUE HANDLER FOR NOT FOUND SET v_done = 1;
    /* Evaluate current situation for every single relevant column and table */
    SET @dynamic_sql = "UPDATE `device_details` SET `MostRecentSecurityEvaluationId` = ?, `CountedSecurityEvaluations` = ?, `LastSeen` = `LastSeen` WHERE (`DeviceId` = ?);";
    PREPARE complete_sql FROM @dynamic_sql;
    OPEN info_cursor;
    REPEAT
        FETCH info_cursor INTO v_DeviceId, v_SecurityEvaluationId, v_CountedSecurityEvaluations;
        IF NOT v_done THEN
            SET @DeviceId = v_DeviceId;
            SET @SecurityEvaluationId = v_SecurityEvaluationId;
            SET @CountedSecurityEvaluations = v_CountedSecurityEvaluations;
            EXECUTE complete_sql USING @SecurityEvaluationId, @CountedSecurityEvaluations, @DeviceId;
        END IF;
    UNTIL v_done END REPEAT;
    CLOSE info_cursor;
    DEALLOCATE PREPARE complete_sql;
END//
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;
