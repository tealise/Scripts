:: Joshua Nasiatka
:: Feb 2016

@echo off
:do_the_things
echo Running script

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '.\winCrypt.ps1'"

:end_all_the_things
exit
