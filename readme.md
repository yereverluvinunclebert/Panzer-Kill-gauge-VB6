# Panzer-Kill-gauge-VB6

A FOSS Kill Gauge VB6 Widget for Reactos, XP, Win7, 8 and 10+.

The Panzer Kill Gauge VB6 gauge is a simple utility displaying a process kill Gauge in a
dieselpunk fashion on your desktop. It can be resized and placed anywhere on
your desktop. It uses minimal resources and helps give your desktop a
dieselpunk make-over.

I keep it in very small size, top right on my screen, ready for use.

I use it to kill a wayward process. I used to run a lot of Yahoo Widgets and
occasionally the Konfabulator engine would get itself in a twist and start
continually dying/reloading. I had this exact widget on my desktop (written
in MS Jscript for the Xwidget engine) and whenever YWE did a 'funny' I would
dblClick on the Kill button and YWE would die. When you are developing it is
not unsusual to create a rogue program that needs to be killed regularly. Use
this tool to do it.

This VB6 version of the kill widget replaces the Xwidget version that is now obsolete.

![killButtonWidget](https://github.com/yereverluvinunclebert/Panzer-Kill-gauge-VB6/assets/2788342/a2110157-3ff9-4d7e-a0e5-fc7b079e4647)

In amongst the code is a binary that you will be able to use to install the
desktop widget. It is SETUP.EXE, when the program is ready I will create this
installer, until then this a work in progress.

![about](https://github.com/yereverluvinunclebert/Panzer-Kill-gauge-VB6/assets/2788342/a91e4fe4-ee6f-4f3a-8b75-9584b60012b0)

This widgets UI functionality is limited as I am using my simple template for
widgets yet to come. However, it can be increased in size, opacity/transparency
may be set as to the users discretion. The widget can also be made to hide for
a pre-determined period.

![panzer-kill-help-750](https://github.com/yereverluvinunclebert/Panzer-Kill-gauge-VB6/assets/2788342/11505f09-61ff-48c9-9627-ed10fa7975ef)

Right-click on the widget to display the function menu, mouse hover over the
widget and press CTRL+mousewheel up/down to resize. It works well on Windows XP
to Windows 11 as well as Linux and Mac os x with Wine.

The Panzer Kill Gauge VB6 gauge is Beta-grade software, under development, not yet
ready to use on a production system - use at your own risk.

![killbuttonmenu01](https://github.com/yereverluvinunclebert/Panzer-Kill-gauge-VB6/assets/2788342/48db2772-feef-4789-a3db-0b217042fcd3)

This version was developed on Windows 7 using 32 bit VisualBasic 6 as a FOSS
project creating a WoW64 widget for the desktop.

![Licence](https://github.com/yereverluvinunclebert/Panzer-Kill-gauge-VB6/assets/2788342/db31cc9d-8fa7-457e-9e76-529b0371432c)

It is open source to allow easy configuration, bug-fixing, enhancement and
community contribution towards free-and-useful VB6 utilities that can be created
by anyone. The first step was the creation of this template program to form the
basis for the conversion of other desktop utilities or widgets. A future step
is new VB6 widgets with more functionality and then hopefully, conversion of
each to RADBasic/TwinBasic for future-proofing and 64bit-ness.

This utility is one of a set of steampunk and dieselpunk widgets. That you can
find here on Deviantart: https://www.deviantart.com/yereverluvinuncleber/gallery

I do hope you enjoy using this utility and others. Your own software
enhancements and contributions will be gratefully received if you choose to
contribute.

![prefsGeneral02](https://github.com/yereverluvinunclebert/Panzer-Kill-gauge-VB6/assets/2788342/2e724c8c-c243-4b9a-95ed-4b43154a1a93)

CREDITS:

I have really tried to maintain the credits as the project has progressed. If I
have made a mistake and left someone out then do forgive me. I will make amends
if anyone points out my mistake in leaving someone out.

MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and
early 2000s. Thanks for VB6.

Olaf Schmidt - This tool was built using the RichClient RC5 Cairo wrapper for
VB6. Specifically the components using transparency. Thanks for the massive
effort Olaf in creating Cairo counterparts for all VB6 native controls and
giving us access to advanced features on controls such as transparency.

Shuja Ali @ codeguru for his settings.ini code.

ALLAPI.COM For the registry reading code.

Rxbagain on codeguru for his Open File common dialog code without a dependent
OCX - http://forums.codeguru.com/member.php?92278-rxbagain

si_the_geek for his special folder code

Elroy on VB forums for the balloon tooltips

Harry Whitfield for his quality testing, brain stimulation and being an
unwitting source of inspiration.

Tested on (with varying results, mostly good!) :

    ReactOS 0.4.14 32bit on virtualBox
    windows 7 professional 32bit on intel
    windows 7 ultimate 64bit on intel
    windows 7 professional 64bit on intel
    windows xp sp3 32bit on intel
    windows 10 home 64bit on intel
    windows 10 home 64bit on amd
    windows 11 64bit on intel
    Open Suse with Wine
    Mac os/x 10.14. Mojave with Wine

Dependencies:

o A windows-alike o/s such as Windows XP, 7-11, Apple Mac OSX 11 with Wine, Linux with Wine.


o Microsoft VB6 IDE installed with its runtime components. The program runs without any additional Microsoft OCX components, just the basic controls that ship with VB6.

![vb6-logo](https://github.com/yereverluvinunclebert/Panzer-Kill-gauge-VB6/assets/2788342/b5aa8d72-0171-48f5-bce6-d74f6a29c7d5)

Requires a PzKill Gauge folder in C:\Users\<user>\AppData\Roaming\
eg: C:\Users\<user>\AppData\Roaming\PzKill Gauge
Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\PzKill Gauge
The above will be created automatically by the compiled program when run for the
first time.

o Krool's replacement for the Microsoft Windows Common Controls found in
mscomctl.ocx (slider) are replicated by the addition of one
dedicated OCX file that are shipped with this package.

During development only, this must be copied to C:\windows\syswow64 and should be registered.

- CCRSlider.ocx

Register this using regsvr32, ie. in a CMD window with administrator privileges.
	
	c:                          ! set device to boot drive with Windows
	cd \windows\syswow64s	    ! change default folder to syswow64
	regsvr32 CCRSlider.ocx	! register the ocx

This will allow the custom controls to be accessible to the VB6 IDE
at design time and the sliders will function as intended (if this ocx is
not registered correctly then the relevant controls will be replaced by picture boxes).

The above is only for development, for ordinary users, during runtime there is no 
need to do the above. The OCX will reside in the program folder. The program reference 
to this OCX is contained within the supplied resource file, saturn.RES. The reference 
to this file is already compiled into the binary. As long as the OCX is in the same 
folder as the binary the program will run without the need to register the OCX manually.

- OLEGuids.tlb

This is a type library that defines types, object interfaces, and more specific
API definitions needed for COM interop / marshalling. It is only used at design
time (IDE). This is a Krool-modified version of the original .tlb from the
vbaccelerator website. The .tlb is compiled into the executable.
For the compiled .exe this is NOT a dependency, only during design time.

From the command line, copy the tlb to a central location (system32 or wow64
folder) and register it.

COPY OLEGUIDS.TLB %SystemRoot%\System32\
REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB

In the VB6 IDE - project - references - browse - select the OLEGuids.tlb

- Uses the RC5 Cairo wrapper from Olaf Schmidt.

During development the RC5 components need to be registered. These scripts are
used to register. Run each by double-clicking on them.

RegisterRC5inPlace.vbs
RegisterVBWidgetsInPlace.vbs

During runtime on the users system, the RC5 components are dynamically
referenced using modRC5regfree.bas which is compiled into the binary.

- SETUP.EXE - The program is currently distributed using setup2go, a very useful
  and comprehensive installer program that builds a .exe installer. Youll have to
  find a copy of setup2go on the web as it is now abandonware. Contact me
  directly for a copy. The file "install PzKill Gauge 0.1.0.s2g" is the configuration
  file for setup2go. When you build it will report any errors in the build.

- HELP.CHM - the program documentation is built using the NVU HTML editor and
  compiled using the Microsoft supplied CHM builder tools (HTMLHelp Workshop) and
  the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
  the job admirably. The HTML files exist alongside the compiled CHM file in the
  HELP folder.

Project References:

    VisualBasic for Applications
    VisualBasic Runtime Objects and Procedures
    VisualBasic Objects and Procedures
    OLE Automation
    vbRichClient5

LICENCE AGREEMENTS:

Copyright © 2023 Dean Beedell

In addition to the GNU General Public Licence please be aware that you may use
any of my own imagery in your own creations but commercially only with my
permission. In all other non-commercial cases I require a credit to the
original artist using my name or one of my pseudonyms and a link to my site.
With regard to the commercial use of incorporated images, permission and a
licence would need to be obtained from the original owner and creator, ie. me.

![killButtonFullScreen02](https://github.com/yereverluvinunclebert/Panzer-Kill-gauge-VB6/assets/2788342/30616201-98bc-4c31-b676-00d25ea8243f)

