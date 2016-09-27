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
  KEY `HostName` (`HostName`),
  KEY `Publisher` (`Publisher`),
  KEY `InstallationDate` (`InstallationDate`),
  KEY `SoftwareNameCleaned` (`SoftwareNameCleaned`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
