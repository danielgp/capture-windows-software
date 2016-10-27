ALTER TABLE `in_windows_software_installed` ADD COLUMN `RegistryHive` ENUM('HKEY_LOCAL_MACHINE', 'HKEY_CURRENT_USER', 'Unknown') NOT NULL DEFAULT 'Unknown' AFTER `RegistryKeyTrunk`;
ALTER TABLE `in_windows_software_installed` ADD COLUMN `Bits32Or64` ENUM('0', '32', '64') NOT NULL DEFAULT '0' AFTER `RegistryHive`;
UPDATE `in_windows_software_installed` SET `RegistryHive` = 'Unknown', `Bits32Or64` = '32' WHERE `RegistryKeyTrunk` = 'Microsoft';
UPDATE `in_windows_software_installed` SET `RegistryHive` = 'Unknown', `Bits32Or64` = '64' WHERE `RegistryKeyTrunk` = 'Wow6432Node';
ALTER TABLE `in_windows_software_installed` DROP PRIMARY KEY;
ALTER TABLE `in_windows_software_installed` ADD PRIMARY KEY(`HostName`, `RegistryHive`, `RegistrySubKey`, `Bits32Or64`);
ALTER TABLE `in_windows_software_installed` DROP COLUMN `RegistryKeyTrunk`;