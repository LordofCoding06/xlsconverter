[String] $folder = "Q:\AppTesting\QFTestFrameWork\QFTestDriver\Syrius";
Get-ChildItem -Path $folder | SELECT Attributes, Name, CreationTime | Format-Table -AutoSize;