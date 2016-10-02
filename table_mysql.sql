CREATE TABLE IF NOT EXISTS `device_details` (
  `DeviceId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `DeviceName` VARCHAR(60) NULL,
  `DeviceOSdetails` JSON NULL DEFAULT NULL,
  `DeviceHardwareDetails` JSON NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`DeviceId`),
  UNIQUE INDEX `ndx_dd_DeviceName_UNIQUE` (`DeviceName` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE `in_windows_software_list` (
  `EvaluationTimestamp` datetime NOT NULL,
  `HostName` varchar(64) NOT NULL,
  `PublisherName` varchar(80) DEFAULT NULL,
  `SoftwareName` varchar(80) NOT NULL,
  `FullVersion` varchar(30) DEFAULT NULL,
  `InstallationDate` date DEFAULT NULL,
  `InstallLocation` text DEFAULT NULL,
  `SizeBytes` mediumint(8) unsigned DEFAULT NULL,
  `OtherInfo` JSON NULL DEFAULT NULL,
  `RegistryKeyTrunk` ENUM('Microsoft', 'Wow6432Node') NOT NULL,
  `RegistrySubKey` varchar(100) NOT NULL,
  PRIMARY KEY(`HostName`, `RegistryKeyTrunk`, `RegistrySubKey`)
  KEY `HostName` (`HostName`),
  KEY `PublisherName` (`PublisherName`),
  KEY `SoftwareName` (`SoftwareName`),
  KEY `FullVersion` (`FullVersion`),
  KEY `InstallationDate` (`InstallationDate`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `publisher_details` (
  `PublisherId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `PublisherName` VARCHAR(80) NOT NULL,
  `PublisheMainWebsite` VARCHAR(250) NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`PublisherId`),
  UNIQUE INDEX `ndx_pd_PublisherName_UNIQUE` (`PublisherName` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `software_details` (
  `SoftwareId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `SoftwareName` VARCHAR(80) NOT NULL,
  `SoftwareDescription` TEXT NULL DEFAULT NULL,
  `SoftwareMainWebsite` VARCHAR(250) NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`SoftwareId`),
  UNIQUE INDEX `ndx_sd_SoftwareName_UNIQUE` (`SoftwareName` ASC)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE `version_details` (
    `FullVersion` varchar(30) NOT NULL,
    `FullVersionParts` JSON GENERATED ALWAYS AS (CONCAT('{ "Major": ', (CASE WHEN (`FullVersion` IS NULL) THEN NULL ELSE CAST(REPLACE((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN `FullVersion` ELSE SUBSTRING_INDEX(`FullVersion`, '.', 1) END), "v", "") AS UNSIGNED) END), ', "Minor": ', (CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 2), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 1), '.'), '') END) AS UNSIGNED) END), ', "Build": ',(CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`FullVersion`) - CHAR_LENGTH(REPLACE(`FullVersion`, '.', ''))) < 2) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 3), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 2), '.'), '') END) AS UNSIGNED) END), ', "Revision": ', (CASE WHEN (`FullVersion` IS NULL) THEN 0 ELSE CAST((CASE WHEN (LOCATE(".", `FullVersion`) = 0) THEN 0 WHEN ((CHAR_LENGTH(`FullVersion`) - CHAR_LENGTH(REPLACE(`FullVersion`, '.', ''))) < 3) THEN 0 ELSE REPLACE(SUBSTRING_INDEX(`FullVersion`, '.', 4), CONCAT(SUBSTRING_INDEX(`FullVersion`, '.', 3), '.'), '') END) AS UNSIGNED) END), ' }')) STORED,
    `FullVersionNumeric` BIGINT(20) UNSIGNED ZEROFILL GENERATED ALWAYS AS (CASE WHEN (`FullVersion` IS NULL) THEN NULL ELSE (CAST(JSON_EXTRACT(`FullVersionParts`, '$.Major') AS UNSIGNED) * POW(10, 14) + CAST(JSON_EXTRACT(`FullVersionParts`, '$.Minor') AS UNSIGNED) * POW(10, 10) + CAST(JSON_EXTRACT(`FullVersionParts`, '$.Build') AS UNSIGNED) * POW(10, 5) + CAST(JSON_EXTRACT(`FullVersionParts`, '$.Revision') AS UNSIGNED)) END) STORED,
    PRIMARY KEY(`FullVersion`),
    KEY `FullVersionNumeric` (`FullVersionNumeric`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
