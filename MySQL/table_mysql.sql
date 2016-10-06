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

CREATE TABLE IF NOT EXISTS `in_windows_software_list` (
  `EvaluationTimestamp` datetime NOT NULL,
  `HostName` varchar(64) NOT NULL,
  `PublisherName` varchar(80) DEFAULT NULL,
  `SoftwareName` varchar(80) NOT NULL,
  `FullVersion` varchar(30) DEFAULT NULL,
  `InstallationDate` date DEFAULT NULL,
  `OtherInfo` JSON NULL DEFAULT NULL,
  `RegistryKeyTrunk` ENUM('Microsoft', 'Wow6432Node') NOT NULL,
  `RegistrySubKey` varchar(100) NOT NULL,
  PRIMARY KEY(`HostName`, `RegistryKeyTrunk`, `RegistrySubKey`),
  KEY `HostName` (`HostName`),
  KEY `PublisherName` (`PublisherName`),
  KEY `SoftwareName` (`SoftwareName`),
  KEY `FullVersion` (`FullVersion`),
  KEY `InstallationDate` (`InstallationDate`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS `in_windows_software_portable` (
  `EvaluationTimestamp` datetime NOT NULL,
  `VolumeSerialNumber` varchar(50) NOT NULL,
  `FileNameSearched` varchar(100) NOT NULL,
  `FilePathFound` text NOT NULL,
  `FileNameFound` varchar(100) DEFAULT NULL,
  `FileVersionFound` varchar(30) DEFAULT NULL,
  `FileSizeFound` mediumint(8) unsigned DEFAULT NULL,
  PRIMARY KEY(`VolumeSerialNumber`, `FileNameSearched`, `FilePathFound`)
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

CREATE TABLE IF NOT EXISTS `version_details` (
    `VersionId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
    `FullVersion` varchar(30) NOT NULL,
    `FullVersionParts` JSON GENERATED ALWAYS AS (CONCAT('{ "Major": ', (CASE WHEN (`FullVersion` IS NULL) THEN NULL ELSE CAST(REPLACE((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN `FullVersion` ELSE SUBSTRING_INDEX(`FullVersion`, '.', 1) END), "v", "") AS UNSIGNED) END), ', "Minor": ', (CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 2), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 1), '.'), '') END) AS UNSIGNED) END), ', "Build": ',(CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`FullVersion`) - CHAR_LENGTH(REPLACE(`FullVersion`, '.', ''))) < 2) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 3), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 2), '.'), '') END) AS UNSIGNED) END), ', "Revision": ', (CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`FullVersion`) - CHAR_LENGTH(REPLACE(`FullVersion`, '.', ''))) < 3) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 4), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 3), '.'), '') END) AS UNSIGNED) END), ' }')) STORED,
    `FullVersionNumeric` BIGINT(20) UNSIGNED ZEROFILL GENERATED ALWAYS AS (CASE WHEN (`FullVersion` IS NULL) THEN NULL ELSE (CAST(JSON_EXTRACT(`FullVersionParts`, '$.Major') AS UNSIGNED) * POW(10, 14) + CAST(JSON_EXTRACT(`FullVersionParts`, '$.Minor') AS UNSIGNED) * POW(10, 10) + CAST(JSON_EXTRACT(`FullVersionParts`, '$.Build') AS UNSIGNED) * POW(10, 5) + CAST(JSON_EXTRACT(`FullVersionParts`, '$.Revision') AS UNSIGNED)) END) STORED,
    PRIMARY KEY(`VersionId`),
    UNIQUE INDEX `ndx_vd_FullVersion_UNIQUE` (`FullVersion` ASC),
    KEY `FullVersionNumeric` (`FullVersionNumeric`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

/* Evaluation structures to support traceability for Software */

CREATE TABLE IF NOT EXISTS `evaluation_headers` (
  `EvaluationId` MEDIUMINT(8) UNSIGNED NOT NULL AUTO_INCREMENT,
  `DeviceId` SMALLINT(5) UNSIGNED NOT NULL,
  `EvaluationTimestamp` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `DateOfGatheringTimestamp` DATETIME DEFAULT NULL,
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
    REPLACE(json_extract(`dd`.`DeviceOSdetails`,'$."Caption"'),'"','') AS `Caption`,
    REPLACE(json_extract(`dd`.`DeviceOSdetails`,'$."OS Architecture"'),'"','') AS `OS_Architecture`,
    REPLACE(json_extract(`dd`.`DeviceOSdetails`,'$."Version"'),'"','') AS `Version`,
    ROUND(CAST(REPLACE(json_extract(`dd`.`DeviceOSdetails`,'$."Total Visible Memory [MB]"'),'"','') AS UNSIGNED) / 1024, 0) AS `RAM [GB]`,
    REPLACE(json_extract(`dd`.`DeviceOSdetails`,'$."Current Time Zone Description"'),'"','') AS `CurrentTimeZone`,
    REPLACE(json_extract(`dd`.`DeviceOSdetails`,'$."OS Language Description"'),'"','') AS `OS_Language`,
    REPLACE(json_extract(`dd`.`DeviceOSdetails`,'$."Locale Description"'),'"','') AS `Locale`,
    COUNT(`eh`.`EvaluationId`) AS `Number of Evaluations`, 
    MAX(`eh`.`DateOfGatheringTimestamp`) AS `Most Recent Evaluation Timestamp`, 
    DATEDIFF(NOW(), MAX(`eh`.`DateOfGatheringTimestamp`)) AS `Most Recent Evaluation Aging` 
FROM `device_details` `dd`
    LEFT JOIN `evaluation_headers` `eh` ON `dd`.`DeviceId` = `eh`.`DeviceId`
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
    (CASE WHEN (`sd`.`SoftwareName` IN ('Intel® Processor Graphics', 'Maxx Audio Installer', 'Microsoft redistributable runtime DLLs', 'Microsoft Visual C++ Additional Runtime', 'Microsoft Visual C++ Minimum Runtime', 'Microsoft Visual C++ Redistributable', 'Microsoft Visual Studio Tools for Office Runtime', 'Office Click-to-Run Extensibility Component', 'Office Click-to-Run Licensing Component', 'Office Click-to-Run Localization Component', 'Security Update for Microsoft .NET Framework')) THEN JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') WHEN (`sd`.`SoftwareName` = 'Intel® Management Engine Components') THEN (CASE WHEN (JSON_EXTRACT(`vd`.`FullVersionParts`, '$.Major') = 1) THEN 1 ELSE "non 1" END) ELSE NULL END) AS `RelevantMajorVersion`,
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
SELECT 8 INTO @crtEvaluationIdToRemove;
DELETE FROM `evaluation_lines` WHERE (`EvaluationId` = @crtEvaluationIdToRemove);
UPDATE `device_details` `dd` SET `dd`.`MostRecentEvaluationId` = (SELECT MAX(`eh`.`EvaluationId`) FROM `evaluation_headers` `eh` WHERE (`eh`.`EvaluationId` < @crtEvaluationIdToRemove) AND (`eh`.`DeviceId` = (SELECT `DeviceId` FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove)))), `LastSeen` = `LastSeen` WHERE (`dd`. `DeviceId` = (SELECT `DeviceId` FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove)));
DELETE FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove);
ALTER TABLE `evaluation_headers` AUTO_INCREMENT = 1;
