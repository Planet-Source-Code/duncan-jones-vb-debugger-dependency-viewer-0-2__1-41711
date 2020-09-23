MCL Debugger and dependency walker
==================================
(c) 2002-2003 Merrion Computing Ltd
42 Ailesbury Mews
Ballsbridge
Dublin 4


Purpose:
--------
A tool to allow you to attach a debugger to a process and be notified when debug events (such as a thread starting, a dll unloading etc.) occur and optionally to pause that application on such an event.

Demonstrates:
-------------
Using the windows API to attach to a process and read it's memory

Use:
----
Firstly set the events (if any) that you want the application being debugged to be paused on.
Then select an application to debug - you can either select from the list of already running applications,
or browse for an application to launch under the control of the debugger.
As the debug events occur the details will appear in the bottom pane.  If you have selected to freeze the debugee application then you will have to press the "Continue" menu when you want to let it continue.
The modules list in the right hand pane will be filled as modules load.  To get extra information about a given module (for exampl,e the imports and exports listing) double click on it and a form will open for each running module.



