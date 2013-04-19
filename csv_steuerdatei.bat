@echo off

rem *******************************
rem *** Vier Steuerdateien ********
rem *** D:Huesemann, 2013-04-19 ***
rem *******************************

echo Steuerdatei SUEDTIROL...
perl csv_steuerdatei.pl "z:\Lernstand_VERA6 - 2013\__VIPP\Vera6_Pilot_Steuerdatei_final_SUEDTIROL.csv" "z:\Lernstand_VERA6 - 2013\__VIPP\vera6_2013\vera6_2013_monster_SUEDTIROL.dbf"
echo Steuerdatei DEUTSCH...
perl csv_steuerdatei.pl "z:\Lernstand_VERA6 - 2013\__VIPP\Vera6_Pilot_Steuerdatei_final_DEUTSCH.csv" "z:\Lernstand_VERA6 - 2013\__VIPP\vera6_2013\vera6_2013_monster_DEUTSCH.dbf"
echo Steuerdatei ENGLISCH...
perl csv_steuerdatei.pl "z:\Lernstand_VERA6 - 2013\__VIPP\Vera6_Pilot_Steuerdatei_final_ENGLISCH.csv" "z:\Lernstand_VERA6 - 2013\__VIPP\vera6_2013\vera6_2013_monster_ENGLISCH.dbf"
echo Steuerdatei MATHEMATIK...
perl csv_steuerdatei.pl "z:\Lernstand_VERA6 - 2013\__VIPP\Vera6_Pilot_Steuerdatei_final_MATHEMATIK.csv" "z:\Lernstand_VERA6 - 2013\__VIPP\vera6_2013\vera6_2013_monster_MATHEMATIK.dbf"
echo.
echo __FERTIG__
echo siehe z:\Lernstand_VERA6 - 2013\__VIPP\vera6_2013\
echo.
