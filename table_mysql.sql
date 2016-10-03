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

CREATE TABLE IF NOT EXISTS `publisher_details` (
  `PublisherId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `PublisherName` VARCHAR(80) NOT NULL,
  `PublisherMainWebsite` VARCHAR(250) NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`PublisherId`),
  UNIQUE INDEX `ndx_pd_PublisherName_UNIQUE` (`PublisherName` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE IF NOT EXISTS `software_details` (
  `SoftwareId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `SoftwareName` VARCHAR(80) NOT NULL,
  `SoftwareDescription` TEXT NULL DEFAULT NULL,
  `SoftwareMainWebsite` VARCHAR(250) NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`SoftwareId`),
  UNIQUE INDEX `ndx_sd_SoftwareName_UNIQUE` (`SoftwareName` ASC)
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
    `device_details`.`DeviceId` AS `DeviceId`,
    `device_details`.`DeviceName` AS `DeviceName`,
    REPLACE(json_extract(`device_details`.`DeviceOSdetails`,'$."Caption"'),'"','') AS `Caption`,
    REPLACE(json_extract(`device_details`.`DeviceOSdetails`,'$."OS Architecture"'),'"','') AS `OS_Architecture`,
    REPLACE(json_extract(`device_details`.`DeviceOSdetails`,'$."Version"'),'"','') AS `Version`,
    REPLACE(json_extract(`device_details`.`DeviceOSdetails`,'$."Total Visible Memory [MB]"'),'"','') AS `TotalVisibleMemoryMB`,
    REPLACE(json_extract(`device_details`.`DeviceOSdetails`,'$."Current Time Zone Description"'),'"','') AS `CurrentTimeZone`,
    REPLACE(json_extract(`device_details`.`DeviceOSdetails`,'$."OS Language Description"'),'"','') AS `OS_Language`,
    REPLACE(json_extract(`device_details`.`DeviceOSdetails`,'$."Locale Description"'),'"','') AS `Locale` 
FROM `device_details`;

CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `view__version_assesment` AS  
SELECT
    `pd`.`PublisherName`,
    `sd`.`SoftwareId`,
    `sd`.`SoftwareName`,
    GROUP_CONCAT(DISTINCT CONCAT(`vd`.`FullVersion`, '_', `dd`.`DeviceName`) ORDER BY `dd`.`DeviceName` SEPARATOR "     ") AS `Version Details`,
    GROUP_CONCAT(DISTINCT `eh`.`EvaluationId` SEPARATOR "; ") AS `Evaluations`,
    MAX(`vd`.`FullVersionNumeric`) AS `Newest`, 
    MIN(`vd`.`FullVersionNumeric`) AS `Oldest`,
    (CASE WHEN (SUM(CASE WHEN (`eh`.`DeviceId` = 3) THEN `vd`.`FullVersionNumeric` ELSE NULL END) IS NULL) THEN "---" WHEN SUM(CASE WHEN (`eh`.`DeviceId` = 3) THEN `vd`.`FullVersionNumeric` ELSE NULL END) = MAX(`vd`.`FullVersionNumeric`) THEN 'Ok' ELSE 'OLD' END) AS `PC2016`,
    (CASE WHEN (SUM(CASE WHEN (`eh`.`DeviceId` = 1) THEN `vd`.`FullVersionNumeric` ELSE NULL END) IS NULL) THEN "---" WHEN SUM(CASE WHEN (`eh`.`DeviceId` = 1) THEN `vd`.`FullVersionNumeric` ELSE NULL END) = MAX(`vd`.`FullVersionNumeric`) THEN 'Ok' ELSE 'OLD' END) AS `PC2014`,
    (CASE WHEN (SUM(CASE WHEN (`eh`.`DeviceId` = 5) THEN `vd`.`FullVersionNumeric` ELSE NULL END) IS NULL) THEN "---" WHEN SUM(CASE WHEN (`eh`.`DeviceId` = 5) THEN `vd`.`FullVersionNumeric` ELSE NULL END) = MAX(`vd`.`FullVersionNumeric`) THEN 'Ok' ELSE 'OLD' END) AS `PC2011`,
    (CASE WHEN MIN(`vd`.`FullVersionNumeric`) = MAX(`vd`.`FullVersionNumeric`) THEN "Everything up-to-date" ELSE "Differences..." END) AS `Assesment` 
FROM `evaluation_lines` `el`
    INNER JOIN `evaluation_headers` `eh` ON `el`.`EvaluationId` = `eh`.`EvaluationId`
    INNER JOIN `device_details` `dd` ON ((`eh`.`EvaluationId` = `dd`.`MostRecentEvaluationId`) AND (`eh`.`DeviceId` = `dd`.`DeviceId`))
    INNER JOIN `publisher_details` `pd` ON `el`.`PublisherId` = `pd`.`PublisherId`
    INNER JOIN `software_details` `sd` ON `el`.`SoftwareId` = `sd`.`SoftwareId`
    INNER JOIN `version_details` `vd` ON `el`.`VersionId` = `vd`.`VersionId`
GROUP BY `sd`.`SoftwareName`;