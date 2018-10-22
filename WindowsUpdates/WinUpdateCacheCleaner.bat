:: Windows Update Cache Cleaner
:: Joshua Nasiatka
:: 2015

@echo off

echo Removing the Cached Files
echo Getting rid of the old leftovers from the fridge
FOR /D %%p IN ("C:\Windows\SoftwareDistribution\Download\*.*") DO rmdir "%%p" /s /q
echo Cleaned the fridge and took out the trash
echo Chores done.
