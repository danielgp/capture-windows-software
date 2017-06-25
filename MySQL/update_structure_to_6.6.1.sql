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