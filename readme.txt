WIN32 API LIBRARY
=================

A library of API declarations and objects


PROJECT COMPLEXITY		ADVANCED
------------------		--------


INTRODUCTION:
=============

This is a project that was started as a result of some other projects that i was working on, i had a need for accessing the registry a lot and instead of re-writing the API routines for every project, i decided to create a dll containing all the necessary routines.

From this i found that i needed a lot of other API declarations throughout my projects and started compiling a library of some of the most common used API declarations. The aim of this project is to provide one reference in your projects with all the necessary API calls in that one reference.

I hope you find this library useful.


LIBRARY CONTENTS:
=================

Object		Description
======		===========
System		This is the main object of the library, has a reference to each of the other objects 		in this library.

OS		This is a Operating System object (properties only), contains information about the 			operating system running on the PC that the code is executed on

Computer	This is a Computer object (properties only), contains information about the computer 			that the code is executed on, and paths to many of the computers system folders.

User		This is a User object (properties only), contains information about the logged on 			user, and paths to many of the users profile folders.

Progress	This is a progress bar calculator, pass in a progress bar control and optional status 		label and this object calculates the percentage done and updates the progress bar and 		optional label.

EventLog	This object provides you with Windows Event Log functionality, define your error 		codes in the message dll (/MsgDLL/) and add your application errors to the windows 		event log.

Registry	The registry object contains many of the most common registry function, read and 		write to and from both the registry and INI files.

Network		The network module provides the capability to map a network drive on the fly and to 		disconnect it when finished with it.

Tray		Add your application to the windows system tray.

SMTP		The SMTP object is aexactly what it says, allowing you to easily create your own SMPT (Still under	email client or integrate it into your existing application to give them email development)	functionality.

Other Public Modules
====================
Constants	Contain all the DLL's constants and enums
Math		Contains some public Math functions
Memory		ZeroMemory
Strings		Contains some public string manipulation routines
Types		Contains all the DLL's user defined types
Window		Contains some public window routines
Error		Contains some public error handling functions

VERSION HISTORY:
================

1.0.0
=====

First version release

1.0.1
=====

Bug fix on ZeroMemory - caused general protection failure. FIXED
Added Network object - Network drive mapping capabilities

2.0.0
=====

Full re-write.
Enhanced all the routines added some objects for the OS, Computer and User. 
Added SMTP capabilities although this is still under development.
Added a global object to reference the other objects.
Added some string manipulation functions.
Added some math routines.

LEGAL MUMBO-JUMBO:
==================
This project and its contents have been scanned for viruses and package as such.  
HOWEVER, It is always best to rescan anything downloaded before using it. I hold no liability as to the user, misuse or abuse of any information released to the public as to its fitness to
a specific use, or losses obtained by its use...  
Long story short - the project works and causes no harm to a computer, however it is still released "AS IS."  
This project may contain components which are freely distributable through third parties.  Any such component is Copyright by its creator.

Technical Support:
==================
This project is released to the public through Planet Source Code and has no warranties implied or expressed.
To report software bug-related comments, suggestions or other project related items, please feel free to
contact the author through means in which this project is distributed.
