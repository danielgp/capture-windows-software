REM --------------------------------------------------------------------------------------------------------------------
REM Should you ever need to Backup your local MySQL database below command line will help (SSL enabled connection)   ---
REM --------------------------------------------------------------------------------------------------------------------
DEL D:\software_monitor.sql
"D:\www\App\MySQL\5.7.x64\bin\mysqldump.exe" --add-locks --allow-keywords --comments --compress --create-options --databases software_monitor --disable-keys --dump-date --enable-cleartext-plugin --events --extended-insert --flush-logs --host=127.0.0.1 --no-autocommit --password --port=3306 --protocol=TCP --quick --quote-names --result-file=D:\software_monitor.sql --routines --set-charset --single-transaction --ssl-ca=D:\www\Config\MySQL\Certificates57x\mysql-ca-cert.pem --ssl-cert=D:\www\Config\MySQL\Certificates57x\mysql-client-cert.pem --ssl-key=D:\www\Config\MySQL\Certificates57x\mysql-client-key.pem --ssl-mode=VERIFY_CA --tls-version=TLSv1.1 --triggers --tz-utc --user=ssluser