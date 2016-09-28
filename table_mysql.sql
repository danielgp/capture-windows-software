CREATE TABLE `in_windows_software_list` (
  `EvaluationTimestamp` datetime NOT NULL,
  `HostName` varchar(64) NOT NULL,
  `Publisher` varchar(80) NOT NULL,
  `Software` varchar(100) NOT NULL,
  `SoftwareNameCleaned` varchar(80) NOT NULL,
  `InstallLocation` text NOT NULL,
  `InstallationDate` date DEFAULT NULL,
  `SizeBytes` mediumint(8) unsigned NOT NULL,
  `VersionMajorMinor` varchar(15) NOT NULL,
  `FullVersionCleaned` varchar(30) NOT NULL,
  `URLinfoAbout` text DEFAULT NULL,
  `RegistryKeyTrunk` varchar(20) NOT NULL,
  `RegistrySubKey` varchar(100) NOT NULL,
  PRIMARY KEY( `RegistryKeyTrunk`, `RegistrySubKey`),
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