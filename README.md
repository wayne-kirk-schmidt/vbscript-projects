Windows VBscript Projects
=========================
The Windows vbscript projects are to capture simple scripts that both 
teach windows batch details and scripting advice when dealing with vbscript.

But why Windows vbscript? Part of any understanding of an OS is what
are the least common denominators for using tools.

By understanding what you can do with the minimum tools allows you to plan
for richer and larger software projects using more full featured tools

That said, vbscript is capable of performing a great deal, so best to
understand what can be done at the command line, then apply to GUI, etc.

Installing the Scripts
=======================

The scripts are organized in directories, and run from them.
Output, Log files, Error files are saved in directories of your choosing from a config file.

Each project can be installed, and then editing the .cfg file to place your output.
Sample Exacution could be:

script /help - generates a help message on how to use the script
script /v   - runs the script verbosely

Dependencies
============

Operating system that knows how to run a a vbscript file.
This will be run typically by:

CMD> cscript //Nologo scriptame.vbs

Script Names and Purposes
=========================

Scripts and Functions:

    1.  getosversion - a cheerful way to get the OS version
        
To Do List:
===========

* add more scripts
* add script functions to mix and match

License
=======

Copyright 2019 Wayne Kirk Schmidt

Licensed under the GNU GPL License (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    license-name   GNU GPL
    license-url    http://www.gnu.org/licenses/gpl.html

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Support
=======

Feel free to e-mail me with issues to: wayne.kirk.schmidt@gmail.com

If you find an issue, please email me! I will investigate and fix.
I will provide best effort fixes and extend the scripts.
