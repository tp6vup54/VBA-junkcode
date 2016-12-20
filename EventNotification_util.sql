-- Summary all ccca logs
EXECUTE sp_CnCSyncSecurityLogtoCnCDetection
EXECUTE sp_CnCSyncWebSecurityLogtoCnCDetection
EXECUTE sp_CnCSyncPFWLogtoCnCDetection
EXECUTE sp_CnCSyncTDALogtoCnCDetection
EXECUTE sp_CnCSyncNCIELogtoCnCDetection

-- Clear all CCCA logs
TRUNCATE TABLE tb_Cncdetection
TRUNCATE TABLE tb_securitylog
TRUNCATE TABLE tb_websecuritylog
TRUNCATE TABLE tb_personalfirewalllog
TRUNCATE TABLE tb_Network_Content_Inspection_Engine_Log
TRUNCATE TABLE tb_loggeneral
TRUNCATE TABLE tb_logmail
TRUNCATE TABLE tb_logip
UPDATE tb_journalcheckpoint SET watermark=0

TRUNCATE TABLE tb_LogDataLossPrevention
TRUNCATE TABLE tb_LogDataLossPreventionTemplate
update tb_LogDataLossPreventionSetting set last_invoke_daily='2011-01-01 00:00:00.000', last_invoke_hourly='2011-01-01 00:00:00.000'