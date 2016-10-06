/* Publishers known info */
INSERT INTO `publisher_known` (`PublisherName`, `PublisherMainWebsite`, `PublisherExtendedInformation`) 
VALUES('Adobe Systems Incorporated', 'https://www.adobe.com/', '{ "Headquarter Address": { "Street Name": "Park Avenue", "Street No": 345, "Location": "San Jose", "State": "California", "Postal Code": "95110-2704", "Country": "United States", "Landline Phone Number": "408-536-6000", "Fax": "408-537-6000" }, "Websites": { "Company": "https://www.adobe.com/about-adobe.html", "Downloads": "https://www.adobe.com/downloads.html" } }'),
('Dell Inc.', 'http://www.dell.com/', '{ "Headquarter Address": { "Street Name": "Dell Way", "Street No": 1, "Location": "Round Rock", "State": "Texas", "Postal Code": 78682, "Country": "United States", "Landline Phone Number": "866-931-3355", "Landline Phone Number 2": "512-338-4400" }, "Websites": { "Company": "http://www.dell.com/learn/us/en/uscorp1/corp-comm", "Downloads": "https://www.dell.com/download" } }'),
('Intel Corporation', 'http://www.intel.com/', '{ "Headquarter Address": { "Street Name": "Oracle Parkway", "Street No": 500, "Location": "Santa Clara", "State": "California", "Postal Code": "95054-1549", "Country": "United States", "Landline Phone Number": "408-765-8080" }, "Websites": { "Company": "http://www.intel.com/content/www/us/en/company-overview/company-overview.html", "Downloads": "https://downloadcenter.intel.com/" } }'),
('Microsoft Corporation', 'https://www.microsoft.com/', '{ "Headquarter Address": { "Street Name": "Microsoft Way", "Street No": 1, "Location": "Redmond", "State": "Washington", "Postal Code": 98052, "Country": "United States" }, "Websites": { "Company": "https://www.microsoft.com/en-us/about/company", "Downloads": "https://www.microsoft.com/en-us/download" } }'),
('Mozilla', 'https://www.mozilla.org/', '{ "Headquarter Address": { "Street Name": "E. Evelyn Avenue", "Street No": 331, "Location": "Mountain View", "State": "California", "Postal Code": 94041, "Country": "United States" }, "Websites": { "Company": "https://www.mozilla.org/en-US/about/" } }'),
('Oracle Corporation', 'https://www.oracle.com/index.html', '{ "Headquarter Address": { "Street Name": "Oracle Parkway", "Street No": 500, "Location": "Redwood City", "State": "California", "Postal Code": 94065, "Country": "United States", "Landline Phone Number": "650-506-7000", "Landline Phone Toll Free Number": "800-392-2999" }, "Websites": { "Company": "https://www.oracle.com/corporate/index.html", "Downloads": "https://www.oracle.com/downloads/index.html" } }');

INSERT INTO `software_files` (`SoftwareFileName`, `SoftwareFileVersionNumericFirst`, `SoftwareFileVersionNumericLast`, `SoftwareName`, `PublisherName`) 
VALUES('7z.exe', 00000000000000000001, 00099900000000000000, '7-Zip', 'Igor Pavlov'),
('7za.exe', 00000000000000000001, 00099900000000000000, '7-Zip Console', 'Igor Pavlov'),
('mysqld.exe', 00000300190000000000, 00000500010002200000, 'MySQL Server', 'MySQL AB'),
('mysqld.exe', 00000500010002300000, 00000500050000700000, 'MySQL Server', 'Sun Microsystems'),
('mysqld.exe', 00000500050000800000, 00099900000000000000, 'MySQL Server', 'Oracle Corporation'),
('MySQLWorkbench.exe', 00000500000000000000, 00000500000002900000, 'MySQL Workbench CE', 'MySQL AB'),
('MySQLWorkbench.exe', 00000500000003000000, 00000500020003107110, 'MySQL Workbench CE', 'Sun Microsystems'),
('MySQLWorkbench.exe', 00000500020003107115, 00099900000000000000, 'MySQL Workbench CE', 'Oracle Corporation'),
('firefox.exe', 00000000010000000000, 00099900000000000000, 'Mozilla Firefox', 'Mozilla'),
('java.exe', 00000500050000800000, 00000600000113000000, 'Java', 'Sun Microsystems'),
('java.exe', 00000600000115000000, 00099900000000000000, 'Java', 'Oracle Corporation'),
('php.exe', 00000100000000000000, 00000500060001400000, 'PHP', 'Zend Technologies'),
('php.exe', 00000500060001500000, 00099900000000000000, 'PHP', 'Rogue Wave Software'),
('notepad++.exe', 00000100000000000000, 00099900000000000000, 'Notepad++', 'Notepad++ Team'),
('httpd.exe', 00000100000000000000, 00099900000000000000, 'Apache HTTP Server', 'Apache Software Foundation'),
('tomcat*w.exe', 00000100000000000000, 00099900000000000000, 'Apache Tomcat', 'Apache Software Foundation'),
('php_xdebug-*.dll', 00000100000000000000, 00099900000000000000, 'X-Debug', 'Derick Rethans'),
('iconv.exe', 00000100000000000000, 00099900000000000000, 'LibIconv', 'Free Software Foundation'),
('msgconv.exe', 00000100000000000000, 00099900000000000000, 'GetText', 'Free Software Foundation'),
('netbeans64.exe', 00000100000000000000, 00000600090000100000, 'NetBeans IDE', 'Sun Microsystems'),
('netbeans64.exe', 00000700000000000000, 00099900000000000000, 'NetBeans IDE', 'Oracle Corporation'),
('ruby.exe', 00000100000000000000, 00099900000000000000, 'Ruby', 'Yukihiro Matsumoto');
