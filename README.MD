OutlookTimeSaver
==============

General Purpose
---------------------
The general purpose of OutlookTimeSaver is to save time writing salutations in outlook 
and to help with small features.

I was stuck with writing every day the same salutation on and on. Also I do have some
customers with complicated and long names, which took me more time checking that I wrote
the name correctly then writing it. In the average I write 20 to 30 emails a day. 
Assuming that I am working with outlook for about 10 years and also long time in future, 
this is a huge mail count. If I only save 5 seconds per email, this would be 10 minutes a week and 
about 7 hours in a year I gain (almost a complete work day).

Also I am tapping me on making the emails in my tray-folder "unread" several times a day. 
This is now automated with this addin. Another time-saver for me.

Copyright
---------------------
Copyright (C) 2016  Jens Hantschel

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.

Why open source?
---------------------
In the beginning with my idea I had the thought of making money with this outlook addin.
I saw that there are several companies selling similar tools (not always based on outlook) and that this
could be possible. After a period of time my mind changed and I became skeptic if this would be rentable.
On the other hand I want that others can use my software.

I came to the following conclusion:

In our times the most valuable thing is time. So I want to make a present for those who need 
time and for those who want to optimise themselve. Hopefully you will succeed.

Feature-List
---------------------
* Automatically sets a default salutation for the recipient or recipients.
** The first- and last-name are read out and the salutation will be written.
* The last salutation to a recipient is saved in a database and will be used next time you write a email to the recipient
* If you move a unread email (bold marked) to your tray it will be marked as "read"
** You won't have to right-click on tray and make all emails unread any more.

Installation
---------------------
You will have to download the Setup-Folder and run the "setup.exe".
The addin will be installed into Outlook and activate on the next start of Outlook.
Please tell me which Outlook-Versions you are using, so that I can make a "compatible list".

Configuration
---------------------
All configuration files and the database with the saved salutation is stored 
under \%appdata%\OutlookTimeSaver\.

The following parameters can configured in OutlookTimeSaverConfig.json:

* DebugViewMode
** Opens a new window with outlook which shows the log entries. This is for debugging purposes
* LoggingEnabled
** Logfiles will be written to %appdata%\OutlookTimeSaver\Log\ if this is enabled
* UnReadDeletedItems
** If this is set to true, deleted items in your tray will be made "unread".
* NoSalutationAtTopicStartsWith
** You can configure different topics where you don't want a automatic salutation. E.g.: This can be 
emails which are generated with a batch file or powershell-script.

Developement
---------------------
OutlookTimeSaver is an addin for Microsoft Outlook.
This addin is developed with Visual Studio 2015 Community Edition.
All you have to do is clone the repository to your computer and 
open it with Visual Studio 2015. After this you will have to download
the nuget packages and start developing.

Planned Features
---------------------
* Translate comments and log-entries from german to englisch, also make future commits in englisch language
** Sorry, in the beginning I did not plan to make this open source.
* Customizable Default-Salutation
* Maybe Multilanguage-Support
* Recognize email formats from previous contacts of same company
** eg. emails with no dot and one letter for firstname --> hmustermann@company.de (Mustermann will be the lastname)
* Statistics how much typing and time is saved
