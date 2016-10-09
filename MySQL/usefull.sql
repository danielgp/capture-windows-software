/*--------------------------------------------------------------------------------------------------------------------*/
/* Should you ever need to remove 1 particular evaluation from the pool use below queries sequence                    */
/*--------------------------------------------------------------------------------------------------------------------*/
SELECT 8 INTO @crtEvaluationIdToRemove;
DELETE FROM `evaluation_lines` WHERE (`EvaluationId` = @crtEvaluationIdToRemove);
UPDATE `device_details` `dd` SET `dd`.`MostRecentEvaluationId` = (SELECT MAX(`eh`.`EvaluationId`) FROM `evaluation_headers` `eh` WHERE (`eh`.`EvaluationId` < @crtEvaluationIdToRemove) AND (`eh`.`DeviceId` = (SELECT `DeviceId` FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove)))), `LastSeen` = `LastSeen` WHERE (`dd`. `DeviceId` = (SELECT `DeviceId` FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove)));
DELETE FROM `evaluation_headers` WHERE (`EvaluationId` = @crtEvaluationIdToRemove);
ALTER TABLE `evaluation_headers` AUTO_INCREMENT = 1;
