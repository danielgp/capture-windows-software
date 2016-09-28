CREATE TABLE `in_windows_software_list` (
  `EvaluationTimestamp` datetime NOT NULL,
  `HostName` varchar(64) NOT NULL,
  `Publisher` varchar(80) DEFAULT NULL,
  `Software` varchar(100) NOT NULL,
  `SoftwareNameCleaned` varchar(80) NOT NULL,
  `InstallLocation` text DEFAULT NULL,
  `InstallationDate` date DEFAULT NULL,
  `SizeBytes` mediumint(8) unsigned DEFAULT NULL,
  `VersionMajorMinor` varchar(15) DEFAULT NULL,
  `FullVersionCleaned` varchar(30) DEFAULT NULL,
  `URLinfoAbout` text DEFAULT NULL,
  `RegistryKeyTrunk` ENUM('Microsoft', 'Wow6432Node') NOT NULL,
  `RegistrySubKey` varchar(100) NOT NULL,
  PRIMARY KEY(`HostName`, `RegistryKeyTrunk`, `RegistrySubKey`)
  KEY `HostName` (`HostName`),
  KEY `Publisher` (`Publisher`),
  KEY `InstallationDate` (`InstallationDate`),
  KEY `SoftwareNameCleaned` (`SoftwareNameCleaned`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `software_monitoring`.`device_details` (
  `DeviceId` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT,
  `DeviceName` VARCHAR(60) NULL,
  `DeviceOSdetails` JSON NULL DEFAULT NULL,
  `DeviceHardwareDetails` JSON NULL DEFAULT NULL,
  `FirstSeen` DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
  `LastSeen` DATETIME(6) DEFAULT NULL ON UPDATE CURRENT_TIMESTAMP(6),
  PRIMARY KEY (`DeviceId`),
  UNIQUE INDEX `ndx_dd_DeviceName_UNIQUE` (`DeviceName` ASC)
) ENGINE = InnoDB DEFAULT CHARSET=utf8;