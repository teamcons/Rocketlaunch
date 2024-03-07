
# -Rocketlaunch!

A program for our job, to create new projects, in the right place, super super quick and how we want it, because we are lazy and impatient.
On the technical aspect, it is a Powershell script, written in imperative style, using Windows.Forms as GUI and compiled into PS2EXE. 
Its bound to have some ugly and some clever code.

It does :
Load the last few emails with attachments from Outlook
Load folder structure from templates
Create in the appropriate structure the folders,
include source files from email with attachments, proper place and naming conventions.

As we charge by the word, count words in Word, Powerpoint, PDF, TXT, create an analysis Excel file and copy total to clipboard
As we work in Trados Studio, pre-fill and start the New Project assistant

<div align="center">
    <img src="https://github.com/teamcons/Skrivanek-Rocketlaunch/blob/main/images/Screenshot App.png" /></td>
</div>


# TODO
-Choose whether to get source files from emails, Downloads folder, or None. Listview for the two latters, with icons
-Load for real from a CSV, and a button to load others. Also icons.
-Drag N Drop



# BUILD

PS2EXE is required if you want to use an EXE and not the PS1 script.
You can install it by opening a powershell window, and entering the command "Install-Module ps2exe"

The folder "Release" has a build script that takes care of using PS2EXE to bundle the script into a nice looking EXE file.

If there is no EXE in that folder, do right-click on "build.ps1", "Execute with Powershell" or whatever it is in your language, and it will take care of generating one for you.


# INSTALL

None required. Just run the EXE in "release" or PS script in "src"
Because of the way this is built, your antivirus may be unhappy. The Windows Defender thingy also sometimes moves away the exe into some quarantine folder.


# Super Skrivanek Suite

This is part of a suite of scripts we coded for our workplace.
We do a nontech job, with a lot of repetitive tasks, and went on to build utilities to automatize that shit.
We arent coders, so the code probably isnt the best, just learning Powershell to make our everyday easier.

The company is Skrivanek GmbH a translation agency, we're there as Project Manager.
The manual is for coworkers who may want to use it.



# Some more stuff

The ability to do rad EXE files is thanks to:
https://github.com/MScholtes/PS2EXE

The rocket icon comes from there:
<a href="https://www.flaticon.com/free-icons/rocket" title="rocket icons">Rocket icons created by Freepik - Flaticon</a>
