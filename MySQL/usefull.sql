/*--------------------------------------------------------------------------------------------------------------------*/
/* Should you ever need to remove 1 particular evaluation from the pool use below queries sequence                    */
/*--------------------------------------------------------------------------------------------------------------------*/
SELECT 8 INTO @crtEvaluationIdToRemove;
DELETE FROM `evaluation_lines` WHERE (`EvaluationId` = @crtEvaluationIdToRemove);
UPDATE `device_details` `dd` SET `dd`.`MostRecentEvaluationId` = (SELECT MAX(`eh`.`EvaluationId`) FROM `evaluation_headers` `eh` WHERE (`eh`.`EvaluationId` < @crtEvaluationIdToRemove) AND (`eh`.`DeviceId` = (SELECT `DeviceId` FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove)))), `LastSeen` = `LastSeen` WHERE (`dd`. `DeviceId` = (SELECT `DeviceId` FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove)));
DELETE FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove);
ALTER TABLE `evaluation_headers` AUTO_INCREMENT = 1;
/*--------------------------------------------------------------------------------------------------------------------*/
/* Should you ever need to remove 1 particular security evaluation from the pool use below queries sequence                    */
/*--------------------------------------------------------------------------------------------------------------------*/
SELECT 8 INTO @crtSecurityEvaluationIdToRemove;
DELETE FROM `security_evaluation_lines` WHERE (`SecurityEvaluationId` = @crtSecurityEvaluationIdToRemove);
UPDATE `device_details` `dd` SET `dd`.`MostRecentSecurityEvaluationId` = (SELECT MAX(`seh`.`SecurityEvaluationId`) FROM `security_evaluation_headers` `seh` WHERE (`seh`.`SecurityEvaluationId` < @crtSecurityEvaluationIdToRemove) AND (`seh`.`DeviceId` = (SELECT `DeviceId` FROM `security_evaluation_headers` WHERE (`SecurityEvaluationId` = @crtSecurityEvaluationIdToRemove)))), `LastSeen` = `LastSeen` WHERE (`dd`. `DeviceId` = (SELECT `DeviceId` FROM `security_evaluation_headers` WHERE (`SecurityEvaluationId` = @crtSecurityEvaluationIdToRemove)));
DELETE FROM `security_evaluation_headers` WHERE (`SecurityEvaluationId` = @crtSecurityEvaluationIdToRemove);
ALTER TABLE `security_evaluation_headers` AUTO_INCREMENT = 1;
