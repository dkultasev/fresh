$settings = Get-Content -LiteralPath "$PSScriptRoot\..\config\settings.json" | ConvertFrom-Json

function Get-ODBC-Data {
    param([string]$query = $(throw 'query is required.'))
    $conn = New-Object System.Data.Odbc.OdbcConnection
    $connStr = "Driver=Firebird/Interbase(r) driver;Server=localhost;Port=3050;Database=$($settings.dbPath);Uid=$($settings.userName);Pwd=$($settings.password);CHARSET=UTF8"
    $conn.ConnectionString = $connStr
    $conn.open
    $cmd = new-object System.Data.Odbc.OdbcCommand($query, $conn)
    $cmd.CommandTimeout = 15
    $ds = New-Object system.Data.DataSet
    $da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
    [void]$da.fill($ds)
    $ds.Tables[0] 
    $conn.close()
}
    
function New-BirthdaysExcel {
    param(
        $Date
    )
    $finalDate = $Date;
    if ($null -eq $finalDate) {
        $finalDate = 'NULL';
    } else {
        $finalDate = "'$finalDate'";
    }

    $query = @"
SELECT
datediff(YEAR,
	birthday,
	THIS_YEAR_BIRTHDAY) years,
CAST(birthday AS date) birthday,
LIST(Card_Num, ', ') Card_Num,
Klient,
contacts
FROM
(
SELECT
	dateadd(datediff (YEAR,
		s.Birthday,
		CURRENT_TIMESTAMP ) YEAR TO s.BIRTHDAY) AS this_year_birthday,
	s.Card_Num,
	s.name AS Klient,
	s.Birthday,
	s.contacts
FROM
	cb_accounts s
WHERE
	s.Birthday IS NOT NULL
	AND CONTACTS IS NOT NULL) a
WHERE
this_year_birthday = COALESCE($finalDate, CAST(CURRENT_TIMESTAMP AS DATE))
GROUP BY datediff(YEAR,
	birthday,
	THIS_YEAR_BIRTHDAY) ,
CAST(birthday AS date) ,
Klient,
contacts
ORDER BY
1 
"@

    $smsTemplate = Get-Content -LiteralPath "$PSScriptRoot\..\template\birthday_template.txt" -Encoding UTF8

    $file = "$PSScriptRoot\..\Birthdays.xlsx";
    if (Test-Path $file ) {
        Remove-Item $file;
    }

    [OfficeOpenXml.ExcelPackage]$excel = New-OOXMLPackage -Author "DK" -Title "Birthdays";
    [OfficeOpenXml.ExcelWorkbook]$book = $excel | Get-OOXMLWorkbook;

    $excel | Add-OOXMLWorksheet -WorkSheetName "Birthdays" #-AutofilterRange "A2:E2"
    $sheet = $book | Select-OOXMLWorkSheet -WorkSheetNumber 1;

    $sheet | Set-OOXMLRangeValue -Row 1 -Col 1 -Value "Age" | Out-Null;
    $sheet | Set-OOXMLRangeValue -Row 1 -Col 2 -Value "Birthday" | Out-Null;
    $sheet | Set-OOXMLRangeValue -Row 1 -Col 3 -Value "Card" | Out-Null;
    $sheet | Set-OOXMLRangeValue -Row 1 -Col 4 -Value "Phone" | Out-Null;
    $sheet | Set-OOXMLRangeValue -Row 1 -Col 5 -Value "Phone2" | Out-Null;
    $sheet | Set-OOXMLRangeValue -Row 1 -Col 6 -Value "Name" | Out-Null;
    $sheet | Set-OOXMLRangeValue -Row 1 -Col 7 -Value "Text" | Out-Null;

    $i = 2;
    $result = Get-ODBC-Data -query $query

    foreach ($row in $result) {
        if (-not($null -eq $row.BIRTHDAY)) {
            $dateStr = $row.BIRTHDAY.ToString("dd.MM.yyyy");
            $phones = $row.CONTACTS.Trim().Split(' ', 2);
            if ($row.CONTACTS.Replace(' ', '').Length -lt 15) {
                $phones[0] = $row.CONTACTS.Replace(' ', '');
                if (-not($null -eq $phones[1])) { $phones[1] = '' };
            }
		
            $names = $row.KLIENT.Split(' ');

            if ($null -eq $names[1]) {
                $firstName = $names[0];
            }
            else {
                $firstName = $names[1];
                $middleName = $names[2];
                ;
            }

            $fullName = "$firstName $middleName".Trim();

            $sheet | Set-OOXMLRangeValue -Row $i -Col 1  -Value $row.years | Out-Null;
            $sheet | Set-OOXMLRangeValue -Row $i -Col 2  -Value $dateStr | Out-Null;
            $sheet | Set-OOXMLRangeValue -Row $i -Col 3  -Value $row.CARD_NUM | Out-Null;
            $sheet | Set-OOXMLRangeValue -Row $i -Col 4  -Value $phones[0].Trim().Replace(' ', '') | Out-Null;
            if ($phones.Count -eq 2) {
                $sheet | Set-OOXMLRangeValue -Row $i -Col 5  -Value $phones[1].Trim().Replace(' ', '') | Out-Null;
            }
            $sheet | Set-OOXMLRangeValue -Row $i -Col 6  -Value $row.KLIENT | Out-Null;
            $sheet | Set-OOXMLRangeValue -Row $i -Col 7  -Value $smsTemplate.Replace('NNNN', $fullName) | Out-Null;
            $i++;
        }
    } 

    $excel | Save-OOXMLPackage -FileFullPath $file -Dispose
}