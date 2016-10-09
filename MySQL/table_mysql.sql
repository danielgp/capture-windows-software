CREATE TABLE IF NOT EXISTS `device_details` (
  `DeviceId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `DeviceName` VARCHAR(60) NULL,
  `DeviceOSdetails` JSON NULL DEFAULT NULL,
  `DeviceHardwareDetails` JSON NULL DEFAULT NULL,
  `MostRecentEvaluationId` MEDIUMINT(8) UNSIGNED DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`DeviceId`),
  UNIQUE INDEX `ndx_dd_DeviceName_UNIQUE` (`DeviceName` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS `device_volumes` (
  `DeviceVolumeId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `VolumeSerialNumber` VARCHAR(60) NULL,
  `DetailedInformation` JSON NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`DeviceVolumeId`),
  UNIQUE INDEX `ndx_dd_VolumeSerialNumber_UNIQUE` (`VolumeSerialNumber` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS `in_windows_software_installed` (
  `EvaluationTimestamp` DATETIME NOT NULL,
  `HostName` VARCHAR(64) NOT NULL,
  `PublisherName` VARCHAR(80) DEFAULT NULL,
  `SoftwareName` VARCHAR(80) NOT NULL,
  `FullVersion` VARCHAR(30) DEFAULT NULL,
  `InstallationDate` date DEFAULT NULL,
  `OtherInfo` JSON NULL DEFAULT NULL,
  `RegistryKeyTrunk` ENUM('Microsoft', 'Wow6432Node') NOT NULL,
  `RegistrySubKey` VARCHAR(100) NOT NULL,
  PRIMARY KEY(`HostName`, `RegistryKeyTrunk`, `RegistrySubKey`),
  KEY `HostName` (`HostName`),
  KEY `PublisherName` (`PublisherName`),
  KEY `SoftwareName` (`SoftwareName`),
  KEY `FullVersion` (`FullVersion`),
  KEY `InstallationDate` (`InstallationDate`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

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
  `FileSizeFound` mediumint(8) unsigned DEFAULT NULL,
  `FilesCheckedForMatchUntilFound` mediumint(8) unsigned NOT NULL,
  PRIMARY KEY(`VolumeSerialNumber`, `FileNameSearched`, `FilePathFound`, `FileNameFound`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS `publisher_details` (
  `PublisherId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `PublisherName` VARCHAR(80) NOT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`PublisherId`),
  UNIQUE INDEX `ndx_pd_PublisherName_UNIQUE` (`PublisherName` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS `publisher_known` (
  `PublisherName` VARCHAR(80) NOT NULL,
  `PublisherMainWebsite` VARCHAR(250) NULL DEFAULT NULL,
  `PublisherExtendedInformation` JSON NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`PublisherName`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS `software_details` (
  `SoftwareId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `SoftwareName` VARCHAR(80) NOT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`SoftwareId`),
  UNIQUE INDEX `ndx_sd_SoftwareName_UNIQUE` (`SoftwareName` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS `software_known` (
  `SoftwareName` VARCHAR(80) NOT NULL,
  `SoftwareDescription` TEXT NULL DEFAULT NULL,
  `SoftwareMainWebsite` VARCHAR(250) NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`SoftwareName`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS `software_files` (
  `SoftwareFileName` VARCHAR(100) NOT NULL,
  `SoftwareFileVersionFirst` VARCHAR(30),
  `SoftwareFileVersionPiecesFirst` JSON GENERATED ALWAYS AS (CONCAT('{ "Major": ', (CASE WHEN (`SoftwareFileVersionFirst` IS NULL) THEN 0 ELSE CAST(REPLACE((CASE WHEN (LOCATE(".", `SoftwareFileVersionFirst`) = 0) THEN `SoftwareFileVersionFirst` ELSE SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 1) END), "v", "") AS UNSIGNED) END), ', "Minor": ', (CASE WHEN (`SoftwareFileVersionFirst` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionFirst`) = 0) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 2), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 1), '.'), '') END) AS UNSIGNED) END), ', "Build": ',(CASE WHEN (`SoftwareFileVersionFirst` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionFirst`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`SoftwareFileVersionFirst`) - CHAR_LENGTH(REPLACE(`SoftwareFileVersionFirst`, '.', ''))) < 2) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 3), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 2), '.'), '') END) AS UNSIGNED) END), ', "Revision": ', (CASE WHEN (`SoftwareFileVersionFirst` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionFirst`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`SoftwareFileVersionFirst`) - CHAR_LENGTH(REPLACE(`SoftwareFileVersionFirst`, '.', ''))) < 3) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 4), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionFirst`, '.', 3), '.'), '') END) AS UNSIGNED) END), ' }')) STORED,
  `SoftwareFileVersionNumericFirst` DECIMAL(25,7) ZEROFILL GENERATED ALWAYS AS (CASE WHEN (`SoftwareFileVersionFirst` IS NULL) THEN NULL ELSE CAST( ( CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesFirst`, '$.Major') * POW(10, 14)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesFirst`, '$.Minor') * POW(10, 7)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesFirst`, '$.Build') * POW(10, 0)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesFirst`, '$.Revision') / POW(10, 7)) AS DECIMAL(8, 7)) ) AS DECIMAL(30, 7)) END) STORED,
  `SoftwareFileVersionLast` VARCHAR(30),
  `SoftwareFileVersionPiecesLast` JSON GENERATED ALWAYS AS (CONCAT('{ "Major": ', (CASE WHEN (`SoftwareFileVersionLast` IS NULL) THEN NULL ELSE CAST(REPLACE((CASE WHEN (LOCATE(".", `SoftwareFileVersionLast`) = 0) THEN `SoftwareFileVersionLast` ELSE SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 1) END), "v", "") AS UNSIGNED) END), ', "Minor": ', (CASE WHEN (`SoftwareFileVersionLast` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionLast`) = 0) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 2), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 1), '.'), '') END) AS UNSIGNED) END), ', "Build": ',(CASE WHEN (`SoftwareFileVersionLast` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionLast`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`SoftwareFileVersionLast`) - CHAR_LENGTH(REPLACE(`SoftwareFileVersionLast`, '.', ''))) < 2) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 3), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 2), '.'), '') END) AS UNSIGNED) END), ', "Revision": ', (CASE WHEN (`SoftwareFileVersionLast` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `SoftwareFileVersionLast`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`SoftwareFileVersionLast`) - CHAR_LENGTH(REPLACE(`SoftwareFileVersionLast`, '.', ''))) < 3) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 4), CONCAT(SUBSTRING_INDEX(`SoftwareFileVersionLast`, '.', 3), '.'), '') END) AS UNSIGNED) END), ' }')) STORED,
  `SoftwareFileVersionNumericLast` DECIMAL(25,7) ZEROFILL GENERATED ALWAYS AS (CASE WHEN (`SoftwareFileVersionLast` IS NULL) THEN NULL ELSE CAST( ( CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesLast`, '$.Major') * POW(10, 14)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesLast`, '$.Minor') * POW(10, 7)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesLast`, '$.Build') * POW(10, 0)) AS UNSIGNED) + CAST((JSON_EXTRACT(`SoftwareFileVersionPiecesLast`, '$.Revision') / POW(10, 7)) AS DECIMAL(8, 7)) ) AS DECIMAL(30, 7)) END) STORED,
  `SoftwareName` VARCHAR(80) NOT NULL,
  `PublisherName` VARCHAR(80) NOT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`SoftwareFileName`, `SoftwareFileVersionFirst`, `SoftwareFileVersionLast`),
  KEY `SoftwareName` (`SoftwareName`),
  KEY `PublisherName` (`PublisherName`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT 'Known files w/ their standardized Software & relevant Publisher';

CREATE TABLE IF NOT EXISTS `version_details` (
    `VersionId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
    `FullVersion` VARCHAR(30) NOT NULL,
    `FullVersionParts` JSON GENERATED ALWAYS AS (CONCAT('{ "Major": ', (CASE WHEN (`FullVersion` IS NULL) THEN NULL ELSE CAST(REPLACE((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN `FullVersion` ELSE SUBSTRING_INDEX(`FullVersion`, '.', 1) END), "v", "") AS UNSIGNED) END), ', "Minor": ', (CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 2), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 1), '.'), '') END) AS UNSIGNED) END), ', "Build": ',(CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`FullVersion`) - CHAR_LENGTH(REPLACE(`FullVersion`, '.', ''))) < 2) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 3), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 2), '.'), '') END) AS UNSIGNED) END), ', "Revision": ', (CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`FullVersion`) - CHAR_LENGTH(REPLACE(`FullVersion`, '.', ''))) < 3) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 4), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 3), '.'), '') END) AS UNSIGNED) END), ' }')) STORED,
    `FullVersionNumeric` DECIMAL(25,7) ZEROFILL GENERATED ALWAYS AS (CASE WHEN (`FullVersion` IS NULL) THEN NULL ELSE CAST( ( CAST((JSON_EXTRACT(`FullVersionParts`, '$.Major') * POW(10, 14)) AS UNSIGNED) + CAST((JSON_EXTRACT(`FullVersionParts`, '$.Minor') * POW(10, 7)) AS UNSIGNED) + CAST((JSON_EXTRACT(`FullVersionParts`, '$.Build') * POW(10, 0)) AS UNSIGNED) + CAST((JSON_EXTRACT(`FullVersionParts`, '$.Revision') / POW(10, 7)) AS DECIMAL(8, 7)) ) AS DECIMAL(30, 7)) END) STORED,
    PRIMARY KEY(`VersionId`),
    UNIQUE INDEX `ndx_vd_FullVersion_UNIQUE` (`FullVersion` ASC),
    KEY `FullVersionNumeric` (`FullVersionNumeric`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS `version_files` (
    `FullVersion` VARCHAR(30) NOT NULL,
    `FullVersionParts` JSON GENERATED ALWAYS AS (CONCAT('{ "Major": ', (CASE WHEN (`FullVersion` IS NULL) THEN NULL ELSE CAST(REPLACE((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN `FullVersion` ELSE SUBSTRING_INDEX(`FullVersion`, '.', 1) END), "v", "") AS UNSIGNED) END), ', "Minor": ', (CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 2), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 1), '.'), '') END) AS UNSIGNED) END), ', "Build": ',(CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`FullVersion`) - CHAR_LENGTH(REPLACE(`FullVersion`, '.', ''))) < 2) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 3), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 2), '.'), '') END) AS UNSIGNED) END), ', "Revision": ', (CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`FullVersion`) - CHAR_LENGTH(REPLACE(`FullVersion`, '.', ''))) < 3) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 4), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 3), '.'), '') END) AS UNSIGNED) END), ' }')) STORED,
    `FullVersionNumeric` DECIMAL(25,7) ZEROFILL GENERATED ALWAYS AS (CASE WHEN (`FullVersion` IS NULL) THEN NULL ELSE CAST( ( CAST((JSON_EXTRACT(`FullVersionParts`, '$.Major') * POW(10, 14)) AS UNSIGNED) + CAST((JSON_EXTRACT(`FullVersionParts`, '$.Minor') * POW(10, 7)) AS UNSIGNED) + CAST((JSON_EXTRACT(`FullVersionParts`, '$.Build') * POW(10, 0)) AS UNSIGNED) + CAST((JSON_EXTRACT(`FullVersionParts`, '$.Revision') / POW(10, 7)) AS DECIMAL(8, 7)) ) AS DECIMAL(30, 7)) END) STORED,
    PRIMARY KEY(`FullVersion`),
    KEY `FullVersionNumeric` (`FullVersionNumeric`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

/* Evaluation structures to support traceability for Software */

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

ALTER TABLE `device_details`
    ADD CONSTRAINT `FK_dd_MostRecentEvaluationId`
        FOREIGN KEY (`MostRecentEvaluationId`)
        REFERENCES `evaluation_headers` (`EvaluationId`)
        ON DELETE RESTRICT
        ON UPDATE CASCADE;

CREATE TABLE IF NOT EXISTS `evaluation_lines` (
  `EvaluationId` MEDIUMINT(8) UNSIGNED NOT NULL,
  `PublisherId` SMALLINT(5) UNSIGNED NOT NULL,
  `SoftwareId` SMALLINT(5) UNSIGNED NOT NULL,
  `VersionId` SMALLINT(5) UNSIGNED NOT NULL,
  `InstallationDate` date DEFAULT NULL,
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
  UNIQUE INDEX `ndx_el_PublisherName_UNIQUE` (`EvaluationId` ASC, `PublisherId` ASC, `SoftwareId` ASC, `VersionId` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

/* View to provide a quick summary on various things */

CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `view__devices` AS
SELECT
    `dd`.`DeviceId` AS `DeviceId`,
    `dd`.`DeviceName` AS `DeviceName`,
    (CASE WHEN (`dd`.`DeviceOSdetails` IS NOT NULL) THEN REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."Caption"'),'"','') ELSE REPLACE(JSON_EXTRACT(`dv`.`DetailedInformation`,'$."Description"'),'"','') END) AS `Caption`,
     (CASE WHEN (`dd`.`DeviceOSdetails` IS NOT NULL) THEN REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."OS Architecture"'),'"','') ELSE REPLACE(JSON_EXTRACT(`dv`.`DetailedInformation`,'$."File System"'),'"','') END) AS `Architecture`,
    REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."Version"'),'"','') AS `Version`,
    ROUND(CAST(REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."Total Visible Memory [MB]"'),'"','') AS UNSIGNED) / 1024, 0) AS `RAM [GB]`,
    REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."Current Time Zone Description"'),'"','') AS `Current Time Zone`,
    REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."OS Language Description"'),'"','') AS `Language`,
    REPLACE(JSON_EXTRACT(`dd`.`DeviceOSdetails`,'$."Locale Description"'),'"','') AS `Locale`,
    COUNT(`eh`.`EvaluationId`) AS `Number of Evaluations`,
    MAX(`eh`.`DateOfGatheringTimestampLast`) AS `Most Recent Evaluation Timestamp`,
    DATEDIFF(NOW(), MAX(`eh`.`DateOfGatheringTimestampLast`)) AS `Most Recent Evaluation Aging`
FROM `device_details` `dd`
    LEFT JOIN `evaluation_headers` `eh` ON `dd`.`DeviceId` = `eh`.`DeviceId`
    LEFT JOIN `device_volumes` `dv` ON `dd`.`DeviceName` = `dv`.`VolumeSerialNumber`
GROUP BY `dd`.`DeviceId`, `dd`.`DeviceName`;

CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `view__evaluations` AS
SELECT
    `el`.`EvaluationId`,
    `dd`.`DeviceId`,
    `dd`.`DeviceName`,
    `pd`.`PublisherId`,
    `pd`.`PublisherName`,
    `sd`.`SoftwareId`,
    `sd`.`SoftwareName`,
    (CASE WHEN (`sd`.`SoftwareName` IN ('Intel® Processor Graphics', 'Maxx Audio Installer', 'Microsoft redistributable runtime DLLs', 'Microsoft Visual C++ Additional Runtime', 'Microsoft Visual C++ Minimum Runtime', 'Microsoft Visual C++ Redistributable', 'Microsoft Visual Studio Tools for Office Runtime', 'Office Click-to-Run Extensibility Component', 'Office Click-to-Run Licensing Component', 'Office Click-to-Run Localization Component', 'PHP', 'Security Update for Microsoft .NET Framework')) THEN JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') WHEN (`sd`.`SoftwareName` = 'Intel® Management Engine Components') THEN (CASE WHEN (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 1) THEN 1 ELSE "non 1" END) WHEN (`sd`.`SoftwareName` IN ('MySQL Documents', 'MySQL Examples and Samples', 'MySQL Server')) THEN (CASE WHEN (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 8) THEN "dmr" ELSE NULL END) WHEN (`sd`.`SoftwareName` IN ('Mozilla Firefox')) THEN (CASE WHEN (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 50) THEN "Beta" WHEN (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 51) THEN "Developer" WHEN (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 52) THEN "Nightly" ELSE NULL END) ELSE NULL END) AS `RelevantMajorVersion`,
    `vd`.`VersionId`,
    `vd`.`FullVersion`,
    `vd`.`FullVersionNumeric`
FROM `evaluation_lines` `el`
    INNER JOIN `evaluation_headers` `eh` ON `el`.`EvaluationId` = `eh`.`EvaluationId`
    INNER JOIN `device_details` `dd` ON ((`eh`.`EvaluationId` = `dd`.`MostRecentEvaluationId`) AND (`eh`.`DeviceId` = `dd`.`DeviceId`))
    INNER JOIN `publisher_details` `pd` ON `el`.`PublisherId` = `pd`.`PublisherId`
    INNER JOIN `software_details` `sd` ON `el`.`SoftwareId` = `sd`.`SoftwareId`
    INNER JOIN `version_details` `vd` ON `el`.`VersionId` = `vd`.`VersionId`;

CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `view__version_assesment` AS
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
HAVING (`Assesment` = 'Differences...');

/* ------------------------------------------------------------------------------------------------------------------ */

DELIMITER //
DROP PROCEDURE IF EXISTS `pr_MatchLatestEvaluationForSoftwarePortrable`//
CREATE PROCEDURE `pr_MatchLatestEvaluationForSoftwarePortrable`()
    NOT DETERMINISTIC
    READS SQL DATA
    SQL SECURITY DEFINER
    COMMENT 'Stores '
BEGIN
    DECLARE v_DeviceId SMALLINT(5) UNSIGNED;
    DECLARE v_EvaluationId MEDIUMINT(8) UNSIGNED;
    DECLARE v_done INT DEFAULT 0;
    /* Reads existing AI columns to later evaluate 1 by 1 */
    DECLARE info_cursor CURSOR FOR SELECT `dd`.`DeviceId`, MAX(`eh`.`EvaluationId`) FROM `in_windows_software_portable` `iwsp` INNER JOIN `device_details` `dd` ON `iwsp`.`VolumeSerialNumber` = `dd`.`DeviceName` INNER JOIN `evaluation_headers` `eh` ON `dd`.`DeviceId` = `eh`.`DeviceId` GROUP BY `dd`.`DeviceId`;
    DECLARE CONTINUE HANDLER FOR NOT FOUND SET v_done = 1;
    /* Evaluate current situation for every single relevant column and table */
    SET @dynamic_sql = "UPDATE `device_details` SET `MostRecentEvaluationId` = ?, `LastSeen` = `LastSeen` WHERE (`DeviceId` = ?);";
    PREPARE complete_sql FROM @dynamic_sql;
    OPEN info_cursor;
    REPEAT
        FETCH info_cursor INTO v_DeviceId, v_EvaluationId;
        IF NOT v_done THEN
            SET @DeviceId = v_DeviceId;
            SET @EvaluationId = v_EvaluationId;
            EXECUTE complete_sql USING @EvaluationId, @DeviceId;
        END IF;
    UNTIL v_done END REPEAT;
    CLOSE info_cursor;
    DEALLOCATE PREPARE complete_sql;
END//
DELIMITER ;

/*--------------------------------------------------------------------------------------------------------------------*/
/* Should you ever need to remove 1 particular evaluation from the pool use below queries sequence                    */
/*--------------------------------------------------------------------------------------------------------------------*/
SELECT 8 INTO @crtEvaluationIdToRemove;
DELETE FROM `evaluation_lines` WHERE (`EvaluationId` = @crtEvaluationIdToRemove);
UPDATE `device_details` `dd` SET `dd`.`MostRecentEvaluationId` = (SELECT MAX(`eh`.`EvaluationId`) FROM `evaluation_headers` `eh` WHERE (`eh`.`EvaluationId` < @crtEvaluationIdToRemove) AND (`eh`.`DeviceId` = (SELECT `DeviceId` FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove)))), `LastSeen` = `LastSeen` WHERE (`dd`. `DeviceId` = (SELECT `DeviceId` FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove)));
DELETE FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove);
ALTER TABLE `evaluation_headers` AUTO_INCREMENT = 1;
