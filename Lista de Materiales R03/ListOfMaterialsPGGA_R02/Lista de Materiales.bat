powershell Expand-Archive ListOfMaterialsPGGA.zip -DestinationPath %USERPROFILE%\Desktop\ -Force
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\Desktop\Lista de Materiales.lnk');$s.TargetPath='%userprofile%\Desktop\ListOfMaterialsPGGA\ListOfMaterialsPGGA\ListOfMaterialsPGGA\bin\Release\ListOfMaterialsPGGA.exe';$s.Save()"
powershell Expand-Archive ListadoDeMateriales_Templates.zip -DestinationPath C:\ABB\ -Force
ren "C:\ABB\PSN_SAS_E3_ARPGA\Template\E3 Export Report Plugin\Template\HPGBillOfMaterial.xlsm" "HPGBillOfMaterial - backup.xlsm"
ren "C:\ABB\PSN_SAS_E3_ARPGA\Template\E3 Export Report Plugin\Template\ABBBillOfMaterial.xlsm" "ABBBillOfMaterial - backup.xlsm"
xcopy ABBBillOfMaterial.xlsm "C:\ABB\PSN_SAS_E3_ARPGA\Template\E3 Export Report Plugin\Template\" /y
xcopy HPGBillOfMaterial.xlsm "C:\ABB\PSN_SAS_E3_ARPGA\Template\E3 Export Report Plugin\Template\" /y
echo "Ya esta listo para usarse el programa!"
pause