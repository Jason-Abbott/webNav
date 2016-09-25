------------------------------------
webNav 2.0
------------------------------------
by Jason Abbott
e-mail: jabbott@uidaho.edu
url: http://boise.uidaho.edu/jason/webNav.html
readme updated: May 20, 1999

webNav 2.0 is an ASP based menuing program distributed in a ZIP file that should include the following files:

webNav2.asp            Main menu view
webNav2_js.asp         JavaScript version of menu
webNav2_choose.html    Determines which menu version to load
webNav2_admin.asp      Form for menu administration
webNav2_edit.asp       Form for editing a menu entry
webNav2_org.asp        Form for re-organizing the menu
webNav2_deleted.asp    Deletes a menu item from the database
webNav2_updated.asp    Updated the database with edited menu entry
webNav2_item.inc       Formats each menu item
webNav2_themes.inc     Included color themes for all pages

data/webNav.mdb        Access 97 database for storing menu items
data/webCal2_array.inc Generates Session array of menu items
data/webCal2_data.inc  Connects to database (creates and opens ADODB object)
webNav2_verify.inc     Checks to see if user has logged in
webNav2_login.asp      Login screen
show_status.inc        Generates javascript to update status bar
images/*.gif           Menu expand and collapse images

Each file contains individual documentation.

INSTALLATION
------------------------------------
Copy the files to a directory beneath the WWW root of your ASP compatible web server.  The name of the main webNav directory is unimportant but the names of the subdirectories /data and /images cannot be changed without updating the menu scripts.  Also, the file names cannot be changed without updating the scripts.

If you would like webNav to choose which menu to load, the JavaScript or non-JavaScript version, then create a link to webNav2_choose.html.  Otherwise, you may create a link directly to either the JavaScript version (webNav2_js.asp) or the non-JavaScript version (webNav2.asp).

You will also need some way to load webNav2_admin.asp.  It's up to you how to load it.  You may create a link to it somewhere, such as in the menu, or you may type it's URL directly.

With the files copied and the links in place, you're done!

ADDING USERS
------------------------------------
The only login account shipped with webNav is username "guest" and password "user."  To add, edit or remove users edit the "users" table in data/webNav.mdb.  A web interface was not developed for this function since many users may choose to use another form of authentication, such as NTFS permissions or IIS settings.

------------------------------------
Thank you for purchasing webNav.  I welcome any questions or feedback that you may have.

Jason Abbott
jabbott@uidaho.edu